using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.IO;
using static VMS.TPS.Common.Model.Types.DoseValue;
using System.Globalization;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.1")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
// [assembly: ESAPIScript(IsWriteable = true)]

namespace GetDataHemangiomas
{
    class Program
    {
        // ======= CONFIG =======
        private const double DVH_BIN_GY = 0.01;

        // Targets gEUD parameter (cold-spot sensitive)
        private const double A_TARGET = -10.0;

        // Brain / Brain-GTV gEUD parameter (typical OAR-like, hot-spot sensitive)
        private const double A_BRAIN = 4.0;

        // Threshold for Brain-GTV Vx
        private const double VX_BRAIN_MINUS_GTV_GY = 32.0;

        // Structure alias maps (Id or Name contains any of these, case-insensitive)
        private static readonly string[] ALIAS_GTV = { "gtv" };
        private static readonly string[] ALIAS_CTV = { "ctv" };
        private static readonly string[] ALIAS_PTV = { "ptv" };

        private static readonly string[] ALIAS_BRAIN = { "brain", "gehirn", "cerebrum" };
        private static readonly string[] ALIAS_BRAINSTEM = { "brainstem", "gehirnstamm", "medulla" };
        private static readonly string[] ALIAS_CHIASM = { "chiasm", "chiasma", "chiasma opticum", "optical chiasm", "opt chiasm" };

        // Actual Brain-GTV structure (e.g., "Brain-GTV")
        private static readonly string[] ALIAS_BRAIN_MINUS_GTV = { "brain-gtv", "brain_minus_gtv", "gehirn-gtv" };

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: ESAPI_PlanStats.exe <inputPatients.csv> <output.csv>");
                    return;
                }

                var inputCsv = args[0];
                var outputCsv = args[1];

                var patientIds = LoadPatientIds(inputCsv);
                if (patientIds.Count == 0)
                {
                    Console.WriteLine("No patient IDs found in input CSV.");
                    return;
                }

                var rows = new List<string>();
                rows.Add(GetHeaderLine());

                using (var app = Application.CreateApplication())
                {
                    foreach (var pid in patientIds)
                    {
                        Console.WriteLine($"Processing PatientID: {pid}");
                        Patient patient = null;
                        try
                        {
                            patient = app.OpenPatientById(pid);
                            if (patient == null)
                            {
                                Console.WriteLine($"  ! Could not open patient {pid} (null). Skipping.");
                                continue;
                            }

                            foreach (var course in patient.Courses ?? Enumerable.Empty<Course>())
                            {
                                foreach (var plan in course.PlanSetups.Where(x => !x.Id.ToUpper()
                                    .Contains("QA(")) ?? Enumerable.Empty<PlanSetup>())
                                {
                                    var line = ProcessPlan(patient, course, plan);
                                    if (line != null)
                                        rows.Add(line);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  ! Error on patient {pid}: {ex.Message}");
                        }
                        finally
                        {
                            if (patient != null)
                                app.ClosePatient();
                        }
                    }
                }

                File.WriteAllLines(outputCsv, rows, Encoding.UTF8);
                Console.WriteLine($"Done. Wrote {rows.Count - 1} plan rows to {outputCsv}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: " + ex);
            }
        }

        private static string GetHeaderLine()
        {
            // Includes Brain/Brain-GTV gEUDs and Brain-GTV V32Gy
            var cols = new[]
            {
                "PatientID","CourseID","PlanID","PrescriptionID","PrescriptionDetails","Anz_Fx",
                "GTV","CTV","PTV",
                "GTV_D50","GTV_D98","GTV_D2",
                "CTV_D50","CTV_D98","CTV_D2",
                "PTV_D50","PTV_D98","PTV_D2",
                "gEUD_GTV","gEUD_CTV","gEUD_PTV",
                "Brain_Volume","Brain_D2","Brain_D50",
                "Brainstem_Volume","Brainstem_D2","Brainstem_D50",
                "Chiasma_Volume","Chiasm_D2","Chiasm_D50",
                "Brain_gEUD_a4","BrainMinusGTV_gEUD_a4","BrainMinusGTV_V32Gy_cm3"
            };
            return string.Join(",", cols);
        }

        private static List<string> LoadPatientIds(string inputCsv)
        {
            var ids = new List<string>();
            foreach (var raw in File.ReadAllLines(inputCsv))
            {
                var line = raw?.Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;

                // Take first non-empty CSV field as patient ID
                var first = line.Split(',').FirstOrDefault()?.Trim();
                if (string.IsNullOrWhiteSpace(first)) continue;

                // Skip header lines
                if (first.ToLowerInvariant().Contains("patient")) continue;

                ids.Add(first);
            }
            return ids.Distinct().ToList();
        }

        private static string ProcessPlan(Patient patient, Course course, PlanSetup plan)
        {
            string pid = patient.Id ?? "";
            string courseId = course.Id ?? "";
            string planId = plan.Id ?? "";

            // Prescription details
            string rxId = plan.RTPrescription?.Id ?? "";
            int? nFx = plan.NumberOfFractions;
            double? dpfGy = plan.DosePerFraction.Dose;
            double? totalDoseGy = plan.TotalDose.Dose;

            if (!dpfGy.HasValue && totalDoseGy.HasValue && nFx.HasValue && nFx.Value > 0)
                dpfGy = totalDoseGy.Value / nFx.Value;

            string rxDetails = BuildRxDetails(dpfGy, nFx, totalDoseGy);

            // Structure set
            var ss = plan.StructureSet;
            if (ss == null)
            {
                return CsvLine(new[]
                {
                    pid, courseId, planId, rxId, rxDetails, nFx?.ToString() ?? "",
                    "", "", "",                          // GTV/CTV/PTV volumes
                    "", "", "", "", "", "", "", "", "",   // DVH targets
                    "", "", "",                          // gEUD targets
                    "", "", "",                          // Brain vol/doses
                    "", "", "",                          // Brainstem vol/doses
                    "", "", "",                          // Chiasm vol/doses
                    "", "", ""                           // Brain gEUD, Brain-GTV gEUD, Brain-GTV V32
                });
            }

            // Find structures
            var (gtv, ctv, ptv) = (
                FindFirst(ss, ALIAS_GTV),
                FindFirst(ss, ALIAS_CTV),
                FindFirst(ss, ALIAS_PTV)
            );

            var brain = FindFirst(ss, ALIAS_BRAIN);
            var brainstem = FindFirst(ss, ALIAS_BRAINSTEM);
            var chiasm = FindFirst(ss, ALIAS_CHIASM);
            var brainMinusGtv = FindFirst(ss, ALIAS_BRAIN_MINUS_GTV);

            // If no dose: just output volumes
            if (plan.Dose == null || !plan.IsDoseValid)
            {
                return CsvLine(new[]
                {
                    pid, courseId, planId, rxId, rxDetails, nFx?.ToString() ?? "",
                    VolStr(gtv), VolStr(ctv), VolStr(ptv),
                    "", "", "", "", "", "", "", "", "",
                    "", "", "",
                    VolStr(brain), "", "",
                    VolStr(brainstem), "", "",
                    VolStr(chiasm), "", "",
                    "", "", ""
                });
            }

            var dvh = new DvhHelper(plan);

            // Targets
            var (gtvD50, gtvD98, gtvD2, gtvGeud) = DoseTripletAndGeud(dvh, gtv, isTarget: true);
            var (ctvD50, ctvD98, ctvD2, ctvGeud) = DoseTripletAndGeud(dvh, ctv, isTarget: true);
            var (ptvD50, ptvD98, ptvD2, ptvGeud) = DoseTripletAndGeud(dvh, ptv, isTarget: true);

            // OAR basics
            var (brainD2, brainD50) = D2D50(dvh, brain);
            var (brainstemD2, brainstemD50) = D2D50(dvh, brainstem);
            var (chiasmD2, chiasmD50) = D2D50(dvh, chiasm);

            // Brain gEUD (a=4)
            double? brain_gEUD_a4 = dvh.gEUD(brain, A_BRAIN);

            // Brain-GTV metrics on actual Brain-GTV structure
            double? brainMinusGtv_gEUD_a4 = dvh.gEUD(brainMinusGtv, A_BRAIN);
            double? brainMinusGtv_V32 = dvh.VxCm3(brainMinusGtv, VX_BRAIN_MINUS_GTV_GY);

            return CsvLine(new[]
            {
                pid, courseId, planId, rxId, rxDetails, nFx?.ToString() ?? "",
                VolStr(gtv), VolStr(ctv), VolStr(ptv),
                DStr(gtvD50), DStr(gtvD98), DStr(gtvD2),
                DStr(ctvD50), DStr(ctvD98), DStr(ctvD2),
                DStr(ptvD50), DStr(ptvD98), DStr(ptvD2),
                DStr(gtvGeud), DStr(ctvGeud), DStr(ptvGeud),
                VolStr(brain), DStr(brainD2), DStr(brainD50),
                VolStr(brainstem), DStr(brainstemD2), DStr(brainstemD50),
                VolStr(chiasm), DStr(chiasmD2), DStr(chiasmD50),
                DStr(brain_gEUD_a4), DStr(brainMinusGtv_gEUD_a4), DStr(brainMinusGtv_V32)
            });
        }

        private static string BuildRxDetails(double? dpfGy, int? nFx, double? totalGy)
        {
            string fx = nFx.HasValue ? nFx.Value.ToString() : "?";
            string dpf = dpfGy.HasValue ? dpfGy.Value.ToString("0.###", CultureInfo.InvariantCulture) : "?";
            string tot = totalGy.HasValue ? totalGy.Value.ToString("0.###", CultureInfo.InvariantCulture) : "?";
            return $"{tot} Gy total; {fx}×{dpf} Gy";
        }

        private static string CsvLine(string[] cells)
            => string.Join(",", cells.Select(EscapeCsv));

        private static string EscapeCsv(string s)
        {
            if (s == null) return "";
            bool needQuotes = s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r");
            var t = s.Replace("\"", "\"\"");
            return needQuotes ? $"\"{t}\"" : t;
        }

        private static string VolStr(Structure s)
            => s == null ? "" : s.Volume.ToString("0.###", CultureInfo.InvariantCulture);

        private static string DStr(double? dGy)
            => dGy.HasValue ? dGy.Value.ToString("0.###", CultureInfo.InvariantCulture) : "";

        private static Structure FindFirst(StructureSet ss, string[] aliases)
        {
            if (ss?.Structures == null) return null;
            foreach (var st in ss.Structures)
            {
                if (st == null || st.IsEmpty) continue;
                var id = (st.Id ?? "").ToLowerInvariant();
                var name = (st.Name ?? "").ToLowerInvariant();
                if (aliases.Any(a => id.Contains(a) || name.Contains(a)))
                    return st;
            }
            return null;
        }

        private static (double? D50, double? D98, double? D2, double? gEUD) DoseTripletAndGeud(
            DvhHelper dvh, Structure s, bool isTarget)
        {
            if (s == null) return (null, null, null, null);
            var d50 = dvh.DoseAtVolumePercent(s, 50.0);
            var d98 = dvh.DoseAtVolumePercent(s, 98.0);
            var d2 = dvh.DoseAtVolumePercent(s, 2.0);
            double? geud = dvh.gEUD(s, A_TARGET); // only used for targets in this code
            return (d50, d98, d2, geud);
        }

        private static (double? D2, double? D50) D2D50(DvhHelper dvh, Structure s)
        {
            if (s == null) return (null, null);
            var d2 = dvh.DoseAtVolumePercent(s, 2.0);
            var d50 = dvh.DoseAtVolumePercent(s, 50.0);
            return (d2, d50);
        }

        // ======= DVH helper =======
        private class DvhHelper
        {
            private readonly PlanSetup _plan;

            public DvhHelper(PlanSetup plan) => _plan = plan;

            public double? DoseAtVolumePercent(Structure s, double volPercent)
            {
                try
                {
                    if (s == null || s.IsEmpty) return null;

                    var dvh = _plan.GetDVHCumulativeData(
                        s,
                        DoseValuePresentation.Absolute,
                        VolumePresentation.Relative,
                        DVH_BIN_GY);   // FIXED

                    if (dvh == null || dvh.CurveData == null || !dvh.CurveData.Any())
                        return null;

                    DVHPoint? prev = null;

                    foreach (var pt in dvh.CurveData)
                    {
                        if (pt.Volume <= volPercent)
                        {
                            // First matching point
                            if (!prev.HasValue)
                                return pt.DoseValue.Dose;

                            // Extract non-nullable values
                            double v1 = prev.Value.Volume;
                            double d1 = prev.Value.DoseValue.Dose;
                            double v2 = pt.Volume;
                            double d2 = pt.DoseValue.Dose;

                            if (Math.Abs(v2 - v1) < 1e-6)
                                return d2;

                            // Linear interpolation
                            double t = (volPercent - v1) / (v2 - v1);
                            return d1 + t * (d2 - d1);
                        }

                        prev = pt;
                    }

                    // If threshold not reached ? highest dose
                    return dvh.CurveData.Last().DoseValue.Dose;
                }
                catch
                {
                    return null;
                }
            }


            // gEUD (Gy) using absolute cm³ DVH
            public double? gEUD(Structure s, double a)
            {
                var (sumPow, vol) = SumDosePowAndVol(s, a);
                if (!sumPow.HasValue || !vol.HasValue || vol.Value <= 0) return null;

                if (a == 0)
                    return Math.Exp(sumPow.Value / vol.Value);

                var meanPow = sumPow.Value / vol.Value;
                return Math.Pow(Math.Max(meanPow, 0.0), 1.0 / a);
            }

            // ?(v_i * d_i^a) and total volume (cm³)
            public (double? sumDosePow, double? totalVol) SumDosePowAndVol(Structure s, double a)
            {
                try
                {
                    if (s == null || s.IsEmpty) return (null, null);

                    var dvh = _plan.GetDVHCumulativeData(
                        s,
                        DoseValuePresentation.Absolute,
                        VolumePresentation.AbsoluteCm3,
                        DVH_BIN_GY);

                    if (dvh == null || dvh.CurveData == null || dvh.CurveData.Count() < 2)
                        return (null, null);

                    var points = dvh.CurveData.ToList();
                    double totalVol = points.First().Volume;
                    if (totalVol <= 0) return (null, null);

                    double sum = 0.0;
                    for (int i = 1; i < points.Count; i++)
                    {
                        var vPrev = points[i - 1].Volume;
                        var vNow = points[i].Volume;
                        var dNow = points[i].DoseValue.Dose;

                        var dv = Math.Max(0.0, vPrev - vNow);
                        if (dv <= 0) continue;

                        sum += dv * Math.Pow(Math.Max(dNow, 0.0), a);
                    }

                    return (sum, totalVol);
                }
                catch
                {
                    return (null, null);
                }
            }

            // Absolute cm³ with dose ? doseGy
            public double? VxCm3(Structure s, double doseGy)
            {
                try
                {
                    if (s == null || s.IsEmpty) return null;

                    var dvh = _plan.GetDVHCumulativeData(
                        s,
                        DoseValuePresentation.Absolute,
                        VolumePresentation.AbsoluteCm3,
                        DVH_BIN_GY);

                    if (dvh == null || dvh.CurveData == null || !dvh.CurveData.Any())
                        return null;

                    DVHPoint? prev = null;

                    foreach (var pt in dvh.CurveData)
                    {
                        if (pt.DoseValue.Dose >= doseGy)
                        {
                            // first point at/above threshold
                            if (!prev.HasValue)
                                return pt.Volume;

                            // use non-nullable doubles for arithmetic
                            double d1 = prev.Value.DoseValue.Dose;
                            double v1 = prev.Value.Volume;
                            double d2 = pt.DoseValue.Dose;
                            double v2 = pt.Volume;

                            if (Math.Abs(d2 - d1) < 1e-6)
                                return v2;

                            double t = (doseGy - d1) / (d2 - d1);
                            return v1 + t * (v2 - v1);
                        }

                        prev = pt;
                    }

                    // threshold above max dose ? 0 cm³
                    return 0.0;
                }
                catch
                {
                    return null;
                }
            }

        }
    }
}
