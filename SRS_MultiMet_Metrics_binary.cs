using System;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// TODO: Uncomment the following line if the script requires write access.
[assembly: ESAPIScript(IsWriteable = true)]
[assembly: AssemblyVersion("1.0.0.13")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]


//forked from https://github.com/Kiragroh/ESAPI_SRS-MultiMets-localMetrics


namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
            // TODO : Add here the code that is called when the script is launched from Eclipse.

            //check if the plan has dose
            if (!context.PlanSetup.IsDoseValid)
            {
                MessageBox.Show("The opened plan has no valid dose.");
                return;
            }

            //get list of structures for loaded plan
            StructureSet ss = context.StructureSet;
            PlanSetup ps = context.PlanSetup;
            var listStructures = ss.Structures.OrderBy(x => x.Id);

            //search for body
            Structure body = listStructures.Where(x => !x.IsEmpty && (x.DicomType.ToUpper().Equals("EXTERNAL") || x.Id.ToUpper().Equals("BODY"))).FirstOrDefault();
            if (body == null)
            {
                MessageBox.Show("Unknown body structure designation. use 'body' or 'Body' (or DicomType 'External').");
                return;
            }

            //enable writing with this script.
            context.Patient.BeginModifications();

            ps.DoseValuePresentation = DoseValuePresentation.Absolute;
            DoseValue d12Gy = new DoseValue(1200, DoseValue.DoseUnit.cGy);
            DoseValue dPrescGy = new DoseValue(ps.TotalDose.Dose, DoseValue.DoseUnit.cGy);

            //define expansion structures, margins, and dummy for local metric calculations
            Structure TargetRing_small_1 = ss.AddStructure("CONTROL", "zTargetRing_small_1");
            TargetRing_small_1.ConvertToHighResolution();
            Structure TargetRing_small_2 = ss.AddStructure("CONTROL", "zTargetRing_small_2");
            TargetRing_small_2.ConvertToHighResolution();
            Structure TargetRing_big_1 = ss.AddStructure("CONTROL", "zTargetRing_big_1");
            TargetRing_big_1.ConvertToHighResolution();
            Structure TargetRing_big_2 = ss.AddStructure("CONTROL", "zTargetRing_big_2");
            TargetRing_big_2.ConvertToHighResolution();
            Structure dummy = ss.AddStructure("CONTROL", "zDummy");

            //calculate TotalMU
            double TotalMU = 0;
            try
            {
                foreach (Beam b in ps.Beams.Where(x => !x.IsSetupField))
                {
                    TotalMU += b.Meterset.Value;
                }
            }
            catch { }

            string msg = string.Format("Local SRS metrics for plan {0} with {1}MU:\n\n", ps.Id, Math.Round(TotalMU, 0));

            //loop through targets
            foreach (Structure target in listStructures.Where(x => !x.IsEmpty && (x.DicomType.ToUpper() == "GTV" || x.DicomType.ToUpper() == "PTV")))
            {
                //check if GTV has a corresponding PTV
                if (target.DicomType.ToUpper() == "GTV")
                {
                    string PTVname = target.Id.ToString().Replace("GTV", "PTV");
                    if (ss.Structures.Any(st => st.Id.Equals(PTVname)))
                    {
                        continue;
                    }
                }

                dPrescGy = new DoseValue(ps.TotalDose.Dose, DoseValue.DoseUnit.cGy);
                //change prescription dose if a totaldose limit for a referencePoint with same ID exists (since Eclipse16 structure names can be longer than 16 chars but RP-Ids not -> therefore this simplification will not always work)
                foreach (ReferencePoint rp in ps.ReferencePoints.Where(x => x.Id == target.Id && x.TotalDoseLimit.ToString() != "N/A"))
                {
                    dPrescGy = new DoseValue(rp.TotalDoseLimit.Dose, DoseValue.DoseUnit.cGy);
                    break;
                }

                //skip targets that have median dose less than prescription
                if (ps.GetDoseAtVolume(target, 50, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose < dPrescGy.Dose)
                {
                    continue;
                }

                //expansions in mm
                double margin_small_1 = 4;  //small expansions used for v100 calc
                double margin_small_2 = 5;
                int margin_big_1 = 9;    //initial expansions for v50 & v12 calc
                int margin_big_2 = 10;
                //expand target
                if (target.IsHighResolution)
                {
                    TargetRing_small_1.SegmentVolume = target.Margin(margin_small_1);
                    TargetRing_small_2.SegmentVolume = target.Margin(margin_small_2);
                    TargetRing_small_2.SegmentVolume = TargetRing_small_1.Or(TargetRing_small_2);
                    TargetRing_big_1.SegmentVolume = target.Margin(margin_big_1);
                    TargetRing_big_2.SegmentVolume = target.Margin(margin_big_2);
                    TargetRing_big_2.SegmentVolume = TargetRing_big_1.Or(TargetRing_big_2);
                }
                else
                {
                    dummy.SegmentVolume = target.SegmentVolume;
                    if (dummy.CanConvertToHighResolution())
                        dummy.ConvertToHighResolution();

                    TargetRing_small_1.SegmentVolume = dummy.Margin(margin_small_1);
                    TargetRing_small_2.SegmentVolume = dummy.Margin(margin_small_2);
                    TargetRing_small_2.SegmentVolume = TargetRing_small_1.Or(TargetRing_small_2);
                    TargetRing_big_1.SegmentVolume = dummy.Margin(margin_big_1);
                    TargetRing_big_2.SegmentVolume = dummy.Margin(margin_big_2);
                    TargetRing_big_2.SegmentVolume = TargetRing_big_1.Or(TargetRing_big_2);
                }

                double v100 = ps.GetVolumeAtDose(TargetRing_small_1, dPrescGy, VolumePresentation.AbsoluteCm3);
                double v100_2 = ps.GetVolumeAtDose(TargetRing_small_2, dPrescGy, VolumePresentation.AbsoluteCm3);
                //Conformity Indices
                double targetVOL = Math.Round(target.Volume, 2);
                double ptv100 = ps.GetVolumeAtDose(target, dPrescGy, VolumePresentation.AbsoluteCm3);
                double pCI = Math.Round((ptv100 * ptv100) / (target.Volume * v100), 2);  //Paddick Conformity Index
                double CI_RTOG = Math.Round(v100 / target.Volume, 2);

                if (Math.Round(v100, 1) != Math.Round(v100_2, 1))
                {
                    pCI = Double.NaN;
                    CI_RTOG = Double.NaN;
                }

                //Gradient Indices & V12  --   increment expansion volume until v50 becomes consistent.  abort at 15mm (presumed bridging)
                double GM = 0.0;
                double GI = 0.0;
                double v12Gy = Double.NaN;
                double v12Gy_1 = 0.0;
                double v12Gy_2 = 0.0;

                for (int i = margin_big_1; i < 15; i++)
                {
                    //local v12
                    v12Gy_1 = Math.Round(ps.GetVolumeAtDose(TargetRing_big_1, d12Gy, VolumePresentation.AbsoluteCm3), 2);
                    v12Gy_2 = Math.Round(ps.GetVolumeAtDose(TargetRing_big_2, d12Gy, VolumePresentation.AbsoluteCm3), 2);
                    //check for dose bridging - this simplification should work based on the assumption that 12Gy is > 50% for all typical SRS Rx's
                    if (Math.Round(v12Gy_1, 1) == Math.Round(v12Gy_2, 1))
                    {
                        v12Gy = v12Gy_2;
                    }

                    double v50 = ps.GetVolumeAtDose(TargetRing_big_1, dPrescGy / 2, VolumePresentation.AbsoluteCm3);
                    double v50_2 = ps.GetVolumeAtDose(TargetRing_big_2, dPrescGy / 2, VolumePresentation.AbsoluteCm3);

                    //Gradient Index
                    GI = Math.Round(v50 / v100, 2);

                    //GradientMeasure  (distance from the radius of sphere matching the volume of V100 to a sphere of V50 volume)
                    double radiusV100 = Math.Pow(((v100 * 3.0 / 4.0) / Math.PI), 1.0 / 3.0);  //must have fractions use decimals to force double instead of integer
                    double radiusV50 = Math.Pow(((v50 * 3.0 / 4.0) / Math.PI), 1.0 / 3.0);
                    GM = Math.Round(radiusV50 - radiusV100, 2);

                    //check for dose bridging. clear the metrics if bridging to prevent misreporting.            
                    if (Math.Round(v50, 1) == Math.Round(v50_2, 1))
                    {
                        break;
                    }
                    else
                    {
                        margin_big_1++;
                        margin_big_2++;

                        TargetRing_big_1.SegmentVolume = target.Margin(margin_big_1);
                        TargetRing_big_2.SegmentVolume = target.Margin(margin_big_2);
                        TargetRing_big_2.SegmentVolume = TargetRing_big_1.Or(TargetRing_big_2);

                        GI = Double.NaN;
                        GM = Double.NaN;
                    }
                }

                //calculate isocenter distance
                Beam firstbeam = ps.Beams.Where(b => b.IsSetupField == false).First();
                VVector isoc = new VVector(Double.NaN, Double.NaN, Double.NaN);
                isoc = firstbeam.IsocenterPosition;
                VVector targetCenter = target.CenterPoint;
                double dist = Math.Round((targetCenter - isoc).Length / 10, 1);

                msg += string.Format("Id: {0}\tVolume: {1}cc \tIsoDist: {2}cm\tDose: {8}cGy\n\tpCI: {3}\tCI_RTOG: {4}\tGI: {5}\tGM: {6}\tlocalV12: {7}cc\n", target.Id, targetVOL, dist, pCI, CI_RTOG, GI, GM, v12Gy, Math.Round(dPrescGy.Dose, 0));
            }

            //delete helper structures
            ss.RemoveStructure(TargetRing_small_1);
            ss.RemoveStructure(TargetRing_small_2);
            ss.RemoveStructure(TargetRing_big_1);
            ss.RemoveStructure(TargetRing_big_2);
            ss.RemoveStructure(dummy);

            //show result
            MessageBox.Show(msg, "SRS Plan Quality Metrics");
        }
    }
}


