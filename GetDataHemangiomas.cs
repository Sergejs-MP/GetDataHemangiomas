using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.IO;
using System.Globalization;

[assembly: AssemblyVersion("1.0.0.1")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]

namespace GetDataHemangiomas
{
    class Program
    {
        // ======= CONFIG =======
        private const double DVH_BIN_GY = 0.01;

        // gEUD parameters
        private const double A_TARGET = -10.0;
        private const double A_BRAIN = 4.0;

        // Brain-GTV Vx threshold
        private const double VX_BRAIN_MINUS_GTV_GY = 32.0;

        // Candidate structure search patterns
        private static readonly string[] ALIAS_TARGETS =
        {
            "gtv", "ctv", "ptv"
        };

        private static readonly string[] ALIAS_BRAIN =
        {
            "brain", "gehirn", "cerebrum"
        };

        private static readonly string[] ALIAS_BRAINSTEM =
        {
            "brainstem", "gehirnstamm", "medulla"
        };

        private static readonly string[] ALIAS_CHIASM =
        {
            "chiasm", "chiasma", "opt chiasm", "chiasma opticum"
        };

        private static readonly string[] ALIAS_BRAIN_MINUS_GTV =
        {
            "brain-gtv", "brain_minus_gtv"
        };

        // ======================
        //        MAIN()
        // ======================

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: GetDataHemangiomas.exe <input.csv> <output.csv>");
                    return;
                }

                string inputCsv = args[0];
                string outputCsv = args[1];

                var triples = LoadPatientCoursePlanTriples(inputCsv);
                if (triples.Count == 0)
                {
                    Console.WriteLine("!! No valid entries in input CSV");
                    return;
                }

                var rows = new List<string>();
                rows.Add(GetHeaderLine());

                using (var app = Application.CreateApplication())
                {
                    foreach (var (pid, courseId, planId) in triples)
                    {
                        Console.WriteLine($"Processing {pid}  {courseId}  {planId}");

                        Patient p = null;

                        try
                        {
                            p = app.OpenPatientById(pid);
                            if (p == null)
                            {
                                Console.WriteLine($"!! Cannot open patient {pid}");
                                continue;
                            }

                            var course = p.Courses
                                .FirstOrDefault(c => c.Id.Equals(courseId, StringComparison.OrdinalIgnoreCase));

                            if (course == null)
                            {
                                Console.WriteLine($"!! Course {courseId} not found for patient {pid}");
                                continue;
                            }

                            var plan = course.PlanSetups
                                .FirstOrDefault(pl => pl.Id.Equals(planId, StringComparison.OrdinalIgnoreCase));

                            if (plan == null)
                            {
                                Console.WriteLine($"!! Plan {planId} not found in course {courseId}");
                                continue;
                            }

                            string row = ProcessPlan(p, course, plan);
                            if (row != null)
                                rows.Add(row);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error on {pid}/{courseId}/{planId}: {ex.Message}");
                        }
                        finally
                        {
                            if (p != null)
                                app.ClosePatient();
                        }
                    }
                }

                File.WriteAllLines(outputCsv, rows, Encoding.UTF8);
                Console.WriteLine($"Done. Exported {rows.Count - 1} plans.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: " + ex.ToString());
            }
        }

        // ======================
        //      CSV LOADER
        // ======================

        private static List<(string pid, string course, string plan)> LoadPatientCoursePlanTriples(string csv)
        {
            var list = new List<(string, string, string)>();

            foreach (var raw in File.ReadAllLines(csv))
            {
                if (string.IsNullOrWhiteSpace(raw)) continue;

                var parts = raw.Split(',');
                if (parts.Length < 3) continue;

                var pid = parts[0].Trim();
                var course = parts[1].Trim();
                var plan = parts[2].Trim();

                // skip header
                if (pid.ToLower().Contains("patient")) continue;

                list.Add((pid, course, plan));
            }

            return list;
        }

        // ======================
        //     HEADER FOR CSV
        // ======================

        private static string GetHeaderLine()
        {
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

        // ======================
        //  PROCESS A SINGLE PLAN
        // ======================

        private static string ProcessPlan(Patient patient, Course course, PlanSetup plan)
        {
            string pid = patient.Id ?? "";
            string courseId = course.Id ?? "";
            string planId = plan.Id ?? "";

            // Prescription
            string rxId = plan.RTPrescription?.Id ?? "";
            int? nFx = plan.NumberOfFractions;
            double? dpf = plan.DosePerFraction.Dose;
            double? total = plan.TotalDose.Dose;

            if (!dpf.HasValue && total.HasValue && nFx.HasValue && nFx > 0)
                dpf = total / nFx;

            string rxDetails = BuildRxDetails(dpf, nFx, total);

            // No structure set => empty row
            if (plan.StructureSet == null)
                return BuildEmptyPlanRow(pid, courseId, planId, rxId, rxDetails, nFx);

            var ss = plan.StructureSet;

            // =========================
            // Find target structures
            // =========================

            // prefix-based expected names

            string prefix = planId.Length >= 2 ? planId.Substring(0, 2) : planId;
            string prefixLong = planId.Length >= 3 ? planId.Substring(0, 3) : planId;


            string gtvName = $"{prefix}_GTV";
            string ctvName = $"{prefix}_CTV";
            string ptvName = $"{prefix}_PTV";

            string gtvNameAlt = $"{prefixLong}_GTV";
            string ctvNameAlt = $"{prefixLong}_CTV";
            string ptvNameAlt = $"{prefixLong}_PTV";

            Structure gtv = FindExactStructure(ss, gtvName);
            if (gtv == null)
                gtv = FindExactStructure(ss, gtvNameAlt);
            Structure ctv = FindExactStructure(ss, ctvName);
            if (ctv == null)
                ctv = FindExactStructure(ss, ctvNameAlt);
            Structure ptv = FindExactStructure(ss, ptvName);
            if (ptv == null)
                ptv = FindExactStructure(ss, ptvNameAlt);

            // OARs
            Structure brain = FindByAlias(ss, ALIAS_BRAIN);
            Structure brainstem = FindByAlias(ss, ALIAS_BRAINSTEM);
            Structure chiasm = FindByAlias(ss, ALIAS_CHIASM);
            Structure brainMinusGtv = FindByAlias(ss, ALIAS_BRAIN_MINUS_GTV);

            bool missingTargets = (gtv == null || ctv == null || ptv == null);

            // DVH unavailable => output volume-only row
            if (plan.Dose == null || !plan.IsDoseValid)
                return BuildDoseMissingRow(pid, courseId, planId, rxId, rxDetails, nFx, gtv, ctv, ptv, brain, brainstem, chiasm);

            var dvh = new DvhHelper(plan);

            // If missing => fallback candidate extraction
            if (missingTargets)
            {
                var candidates = FindCandidateTargets(ss);

                // Write full dosimetry for all candidates
                foreach (var s in candidates)
                {
                    WriteCandidateStructureDose(
                        "CandidateStructures",
                        patient,
                        course,
                        plan,
                        s,
                        dvh
                    );
                }

                // Optional: auto-select best
                if (gtv == null)
                    gtv = candidates.FirstOrDefault(s => s.Id.ToLower().Contains("gtv"));

                if (ctv == null)
                    ctv = candidates
                        .Where(s => s.Id.ToLower().Contains("ctv"))
                        .OrderByDescending(s => s.Volume)
                        .FirstOrDefault();

                if (ptv == null)
                    ptv = candidates
                        .Where(s => s.Id.ToLower().Contains("ptv"))
                        .OrderByDescending(s => s.Volume)
                        .FirstOrDefault();
            }
            // ====== TARGET DOSE METRICS ======
            var (gtvD50, gtvD98, gtvD2, gtvGeud) = DoseTripletAndGeud(dvh, gtv);
            var (ctvD50, ctvD98, ctvD2, ctvGeud) = DoseTripletAndGeud(dvh, ctv);
            var (ptvD50, ptvD98, ptvD2, ptvGeud) = DoseTripletAndGeud(dvh, ptv);

            // ====== OAR DOSE METRICS ======
            var (brainD2, brainD50) = D2D50(dvh, brain);
            var (brainstemD2, brainstemD50) = D2D50(dvh, brainstem);
            var (chiasmD2, chiasmD50) = D2D50(dvh, chiasm);

            double? brain_gEUD_a4 = dvh.gEUD(brain, A_BRAIN);
            double? brainMinusGtv_gEUD_a4 = dvh.gEUD(brainMinusGtv, A_BRAIN);
            double? brainMinusGtv_V32 = dvh.VxCm3(brainMinusGtv, VX_BRAIN_MINUS_GTV_GY);

            // ====== BUILD MAIN CSV ROW ======

            return CsvLine(new[]
            {
                pid, courseId, planId,
                rxId, rxDetails,
                nFx?.ToString() ?? "",

                VolStr(gtv), VolStr(ctv), VolStr(ptv),

                DStr(gtvD50), DStr(gtvD98), DStr(gtvD2),
                DStr(ctvD50), DStr(ctvD98), DStr(ctvD2),
                DStr(ptvD50), DStr(ptvD98), DStr(ptvD2),

                DStr(gtvGeud), DStr(ctvGeud), DStr(ptvGeud),

                VolStr(brain), DStr(brainD2), DStr(brainD50),
                VolStr(brainstem), DStr(brainstemD2), DStr(brainstemD50),
                VolStr(chiasm), DStr(chiasmD2), DStr(chiasmD50),

                DStr(brain_gEUD_a4),
                DStr(brainMinusGtv_gEUD_a4),
                DStr(brainMinusGtv_V32)
            });
        }

        // =======================
        //     FALLBACK EXPORT
        // =======================

        private static List<Structure> FindCandidateTargets(StructureSet ss)
        {
            if (ss == null || ss.Structures == null) return new List<Structure>();

            return ss.Structures
                .Where(s =>
                    !s.IsEmpty &&
                    ALIAS_TARGETS.Any(t =>
                        s.Id.ToLower().Contains(t) ||
                        s.Name.ToLower().Contains(t)))
                .ToList();
        }

        private static void WriteCandidateStructureDose(
            string folder,
            Patient pat,
            Course course,
            PlanSetup plan,
            Structure s,
            DvhHelper dvh)
        {
            Directory.CreateDirectory(folder);

            string file = Path.Combine(folder, $"{pat.Id}_CandidateTargets.csv");

            double? d2 = dvh.DoseAtVolumePercent(s, 2);
            double? d50 = dvh.DoseAtVolumePercent(s, 50);
            double? d98 = dvh.DoseAtVolumePercent(s, 98);
            double? geud = dvh.gEUD(s, A_TARGET);

            using (var sw = new StreamWriter(file, append: true))
            {
                sw.WriteLine(string.Join(",", new[]
                {
                    pat.Id,
                    course.Id,
                    plan.Id,
                    s.Id,
                    s.Name,
                    s.Volume.ToString("0.###", CultureInfo.InvariantCulture),
                    DStr(d2),
                    DStr(d50),
                    DStr(d98),
                    DStr(geud)
                }));
            }
        }

        // =======================
        //   STRUCTURE FINDING
        // =======================

        private static Structure FindExactStructure(StructureSet ss, string name)
        {
            return ss.Structures
                .FirstOrDefault(s =>
                    s.Id.StartsWith(name, StringComparison.OrdinalIgnoreCase) ||
                    s.Name.StartsWith(name, StringComparison.OrdinalIgnoreCase));
        }

        private static Structure FindByAlias(StructureSet ss, string[] aliases)
        {
            foreach (var s in ss.Structures)
            {
                if (s == null || s.IsEmpty) continue;

                string id = (s.Id ?? "").ToLower();
                string nm = (s.Name ?? "").ToLower();

                if (aliases.Any(a => id.Contains(a) || nm.Contains(a)))
                    return s;
            }

            return null;
        }

        // ======================
        //   BUILD EMPTY ROWS
        // ======================

        private static string BuildEmptyPlanRow(
            string pid, string course, string plan,
            string rxId, string rxDetails, int? nFx)
        {
            return string.Join(",", new[]
            {
                pid, course, plan,
                rxId, rxDetails,
                nFx?.ToString() ?? "",
                "", "", "",
                "", "", "", "", "", "", "", "", "",
                "", "", "",
                "", "", "",
                "", "", "",
                "", "", "",
                "", "", ""
            });
        }

        private static string BuildDoseMissingRow(
            string pid, string course, string plan,
            string rxId, string rxDetails, int? nFx,
            Structure gtv, Structure ctv, Structure ptv,
            Structure brain, Structure brainstem, Structure chiasm)
        {
            return string.Join(",", new[]
            {
                pid, course, plan,
                rxId, rxDetails,
                nFx?.ToString() ?? "",
                VolStr(gtv), VolStr(ctv), VolStr(ptv),
                "", "", "", "", "", "", "", "", "",
                "", "", "",
                VolStr(brain), "", "",
                VolStr(brainstem), "", "",
                VolStr(chiasm), "", "",
                "", "", ""
            });
        }

        // ======================
        //   STRING HELPERS
        // ======================

        private static string CsvLine(string[] cells)
        {
            return string.Join(",", cells.Select(EscapeCsv));
        }

        private static string EscapeCsv(string s)
        {
            if (s == null) return "";
            bool mustQuote = s.Contains(",") || s.Contains("\"") || s.Contains("\n");
            string t = s.Replace("\"", "\"\"");
            return mustQuote ? $"\"{t}\"" : t;
        }

        private static string VolStr(Structure s)
        {
            if (s == null) return "";
            return s.Volume.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static string DStr(double? d)
        {
            return d.HasValue
                ? d.Value.ToString("0.###", CultureInfo.InvariantCulture)
                : "";
        }

        private static string BuildRxDetails(double? dpf, int? nFx, double? tot)
        {
            string fx = nFx.HasValue ? nFx.Value.ToString() : "?";
            string dp = dpf.HasValue ? dpf.Value.ToString("0.###", CultureInfo.InvariantCulture) : "?";
            string tt = tot.HasValue ? tot.Value.ToString("0.###", CultureInfo.InvariantCulture) : "?";
            return $"{tt} Gy total; {fx}×{dp} Gy";
        }

        // ======================
        //   DOSE COMPUTATION
        // ======================

        private static (double? D50, double? D98, double? D2, double? gEUD)
            DoseTripletAndGeud(DvhHelper dvh, Structure s)
        {
            if (s == null)
                return (null, null, null, null);

            var d50 = dvh.DoseAtVolumePercent(s, 50);
            var d98 = dvh.DoseAtVolumePercent(s, 98);
            var d2 = dvh.DoseAtVolumePercent(s, 2);
            var ge = dvh.gEUD(s, A_TARGET);

            return (d50, d98, d2, ge);
        }

        private static (double? D2, double? D50) D2D50(DvhHelper dvh, Structure s)
        {
            if (s == null) return (null, null);
            return (
                dvh.DoseAtVolumePercent(s, 2),
                dvh.DoseAtVolumePercent(s, 50)
            );
        }

        // ======================
        //      DVH HELPER
        // ======================

        private class DvhHelper
        {
            private readonly PlanSetup _plan;

            public DvhHelper(PlanSetup plan)
            {
                _plan = plan;
            }

            // ===== DOSE AT VOLUME % =====

            public double? DoseAtVolumePercent(Structure s, double volPercent)
            {
                try
                {
                    if (s == null || s.IsEmpty) return null;

                    var dvh = _plan.GetDVHCumulativeData(
                        s,
                        DoseValuePresentation.Absolute,
                        VolumePresentation.Relative,
                        DVH_BIN_GY);  // correct constructor

                    if (dvh == null || dvh.CurveData == null || !dvh.CurveData.Any())
                        return null;

                    DVHPoint? prev = null;

                    foreach (var pt in dvh.CurveData)
                    {
                        if (pt.Volume <= volPercent)
                        {
                            if (!prev.HasValue)
                                return pt.DoseValue.Dose;

                            double v1 = prev.Value.Volume;
                            double d1 = prev.Value.DoseValue.Dose;

                            double v2 = pt.Volume;
                            double d2 = pt.DoseValue.Dose;

                            if (Math.Abs(v2 - v1) < 1e-6)
                                return d2;

                            double t = (volPercent - v1) / (v2 - v1);
                            return d1 + t * (d2 - d1);
                        }

                        prev = pt;
                    }

                    // if threshold never reached → return smallest volume point’s dose
                    return dvh.CurveData.Last().DoseValue.Dose;
                }
                catch
                {
                    return null;
                }
            }

            // ===== gEUD =====

            public double? gEUD(Structure s, double a)
            {
                var (sumPow, vol) = SumDosePowAndVol(s, a);

                if (!sumPow.HasValue || !vol.HasValue || vol.Value <= 0)
                    return null;

                if (a == 0)
                    return Math.Exp(sumPow.Value / vol.Value);

                double meanPow = sumPow.Value / vol.Value;
                return Math.Pow(Math.Max(meanPow, 0.0), 1.0 / a);
            }

            // ===== DIFFERENTIAL-SUM FORM FOR gEUD =====

            public (double? sumDosePow, double? totalVol) SumDosePowAndVol(Structure s, double a)
            {
                try
                {
                    if (s == null || s.IsEmpty)
                        return (null, null);

                    var dvh = _plan.GetDVHCumulativeData(
                        s,
                        DoseValuePresentation.Absolute,
                        VolumePresentation.AbsoluteCm3,
                       DVH_BIN_GY);

                    if (dvh == null || dvh.CurveData == null || dvh.CurveData.Count() < 2)
                        return (null, null);

                    var points = dvh.CurveData.ToList();
                    double totalVol = points.First().Volume; // cm³

                    if (totalVol <= 0)
                        return (null, null);

                    double sum = 0.0;

                    for (int i = 1; i < points.Count; i++)
                    {
                        double vPrev = points[i - 1].Volume;
                        double vNow = points[i].Volume;
                        double dNow = points[i].DoseValue.Dose;

                        double dv = Math.Max(0.0, vPrev - vNow);
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

            // ===== ABSOLUTE VOLUME ABOVE DOSE =====

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

                    if (dvh == null || dvh.CurveData == null || dvh.CurveData.Count() == 0)
                        return null;

                    DVHPoint? prev = null;

                    foreach (var pt in dvh.CurveData)
                    {
                        if (pt.DoseValue.Dose >= doseGy)
                        {
                            if (!prev.HasValue)
                                return pt.Volume;

                            double d1 = prev.Value.DoseValue.Dose;
                            double d2 = pt.DoseValue.Dose;

                            double v1 = prev.Value.Volume;
                            double v2 = pt.Volume;

                            if (Math.Abs(d2 - d1) < 1e-6)
                                return v2;

                            double t = (doseGy - d1) / (d2 - d1);
                            return v1 + t * (v2 - v1);
                        }

                        prev = pt;
                    }

                    // threshold > max dose → zero
                    return 0.0;
                }
                catch
                {
                    return null;
                }
            }
        } // end of DvhHelper

    } // end class Program
} // end namespace
