using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: AssemblyVersion("1.0.0.1")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]

namespace GetDataHemangiomas
{
    class Program
    {
        private const double DvhBinWidthGy = 0.01;
        private const double MinimumDvhCoverage = 0.999;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                if (args.Length != 3)
                {
                    WriteUsage();
                    Environment.ExitCode = 2;
                    return;
                }

                string inputCsv = Path.GetFullPath(args[0]);
                string roiListFile = Path.GetFullPath(args[1]);
                string outputCsv = Path.GetFullPath(args[2]);

                ValidatePaths(inputCsv, roiListFile, outputCsv);

                var requests = LoadPatientPlanRequests(inputCsv);
                var roiIds = LoadRoiIds(roiListFile);

                if (requests.Count == 0)
                    throw new InvalidDataException("The input CSV contains no patient-plan requests.");

                if (roiIds.Count == 0)
                    throw new InvalidDataException("The ROI list contains no ROI IDs.");

                var rows = new List<string> { BuildHeaderLine(roiIds) };
                string outputDirectory = Path.GetDirectoryName(outputCsv);

                if (!string.IsNullOrEmpty(outputDirectory))
                    Directory.CreateDirectory(outputDirectory);

                bool patientCloseFailed = false;

                using (var app = Application.CreateApplication())
                {
                    foreach (var request in requests)
                    {
                        Console.WriteLine(
                            $"Processing {request.PatientId}  {request.CourseId}  {request.PlanId}");

                        Patient patient = null;
                        PlanExtractionResult result;
                        string closeError = null;

                        if (patientCloseFailed)
                        {
                            result = CreateFailureResult(
                                request,
                                roiIds,
                                "SESSION_ABORTED",
                                "No patient was opened because a previous patient could not be closed safely.");
                            rows.Add(BuildOutputLine(result));
                            continue;
                        }

                        try
                        {
                            patient = app.OpenPatientById(request.PatientId);

                            if (patient == null)
                            {
                                result = CreateFailureResult(
                                    request,
                                    roiIds,
                                    "PATIENT_NOT_FOUND",
                                    "Patient could not be opened.");
                            }
                            else
                            {
                                var course = patient.Courses.FirstOrDefault(c =>
                                    string.Equals(
                                        c.Id,
                                        request.CourseId,
                                        StringComparison.OrdinalIgnoreCase));

                                if (course == null)
                                {
                                    result = CreateFailureResult(
                                        request,
                                        roiIds,
                                        "COURSE_NOT_FOUND",
                                        $"Derived course {request.CourseId} was not found.");
                                }
                                else
                                {
                                    var plan = course.PlanSetups.FirstOrDefault(p =>
                                        string.Equals(
                                            p.Id,
                                            request.PlanId,
                                            StringComparison.OrdinalIgnoreCase));

                                    result = plan == null
                                        ? CreateFailureResult(
                                            request,
                                            roiIds,
                                            "PLAN_NOT_FOUND",
                                            "The exact plan ID was not found in the derived course.")
                                        : ExtractPlan(request, plan, roiIds);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(
                                $"Error on {request.PatientId}/{request.CourseId}/{request.PlanId}: {ex.Message}");

                            result = CreateFailureResult(
                                request,
                                roiIds,
                                "ERROR",
                                ex.Message);
                        }
                        finally
                        {
                            if (patient != null)
                            {
                                try
                                {
                                    app.ClosePatient();
                                }
                                catch (Exception ex)
                                {
                                    closeError = string.IsNullOrWhiteSpace(ex.Message)
                                        ? "Unknown patient-close error."
                                        : ex.Message;
                                    patientCloseFailed = true;
                                    Console.WriteLine(
                                        $"Could not close patient {request.PatientId}: {closeError}");
                                }
                            }
                        }

                        if (closeError != null)
                        {
                            result.PlanStatus = "CLOSE_PATIENT_ERROR";
                            result.Message = AppendMessage(
                                result.Message,
                                "Patient close failed; remaining requests were not opened. " + closeError);
                        }

                        rows.Add(BuildOutputLine(result));
                    }

                    WriteOutputAtomically(outputCsv, rows);
                }

                if (patientCloseFailed)
                {
                    Console.WriteLine(
                        $"Stopped safely after a patient-close error. Wrote {requests.Count} status rows to {outputCsv}.");
                    Environment.ExitCode = 1;
                }
                else
                {
                    Console.WriteLine($"Done. Exported {requests.Count} plan rows to {outputCsv}.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: " + ex.Message);
                Environment.ExitCode = 1;
            }
        }

        private static void WriteUsage()
        {
            Console.WriteLine(
                "Usage: GetDataHemangiomas.exe <patients-plans.csv> <roi-ids.txt> <output.csv>");
            Console.WriteLine("  patients-plans.csv columns: PatientID,PlanID");
            Console.WriteLine("  roi-ids.txt: one exact Eclipse ROI ID per line");
            Console.WriteLine("  CourseID is derived from the first digit in PlanID.");
        }

        private static void ValidatePaths(string inputCsv, string roiListFile, string outputCsv)
        {
            if (!File.Exists(inputCsv))
                throw new FileNotFoundException("Input CSV was not found.", inputCsv);

            if (!File.Exists(roiListFile))
                throw new FileNotFoundException("ROI list was not found.", roiListFile);

            if (PathsEqual(inputCsv, roiListFile))
                throw new InvalidDataException("Input CSV and ROI list must be different files.");

            if (PathsEqual(inputCsv, outputCsv) || PathsEqual(roiListFile, outputCsv))
                throw new InvalidDataException("Output CSV must not overwrite either input file.");
        }

        private static bool PathsEqual(string first, string second)
        {
            return string.Equals(
                Path.GetFullPath(first).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                Path.GetFullPath(second).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                StringComparison.OrdinalIgnoreCase);
        }

        private static List<PatientPlanRequest> LoadPatientPlanRequests(string csvPath)
        {
            var requests = new List<PatientPlanRequest>();
            int patientColumn = 0;
            int planColumn = 1;
            bool firstContentLine = true;

            string[] lines = File.ReadAllLines(csvPath);

            for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
            {
                string raw = lines[lineIndex];
                if (string.IsNullOrWhiteSpace(raw))
                    continue;

                List<string> cells = ParseCsvLine(raw, lineIndex + 1);

                if (firstContentLine)
                {
                    firstContentLine = false;

                    int detectedPatientColumn = FindHeaderColumn(cells, "patientid");
                    int detectedPlanColumn = FindHeaderColumn(cells, "planid");

                    if (detectedPatientColumn >= 0 || detectedPlanColumn >= 0)
                    {
                        if (detectedPatientColumn < 0 || detectedPlanColumn < 0)
                        {
                            throw new InvalidDataException(
                                $"Input CSV header on line {lineIndex + 1} must contain PatientID and PlanID.");
                        }

                        patientColumn = detectedPatientColumn;
                        planColumn = detectedPlanColumn;
                        continue;
                    }
                }

                int requiredColumn = Math.Max(patientColumn, planColumn);
                if (cells.Count <= requiredColumn)
                {
                    throw new InvalidDataException(
                        $"Input CSV line {lineIndex + 1} does not contain both PatientID and PlanID.");
                }

                string patientId = cells[patientColumn].Trim();
                string planId = cells[planColumn].Trim();

                if (patientId.Length == 0 || planId.Length == 0)
                {
                    throw new InvalidDataException(
                        $"Input CSV line {lineIndex + 1} has an empty PatientID or PlanID.");
                }

                string courseId = DeriveCourseId(planId);
                if (courseId == null)
                {
                    throw new InvalidDataException(
                        $"PlanID on input CSV line {lineIndex + 1} contains no digit for CourseID.");
                }

                requests.Add(new PatientPlanRequest(patientId, planId, courseId));
            }

            return requests;
        }

        private static int FindHeaderColumn(IList<string> cells, string expectedHeader)
        {
            for (int i = 0; i < cells.Count; i++)
            {
                string normalized = new string(cells[i]
                    .Where(char.IsLetterOrDigit)
                    .Select(char.ToLowerInvariant)
                    .ToArray());

                if (normalized == expectedHeader)
                    return i;
            }

            return -1;
        }

        private static string DeriveCourseId(string planId)
        {
            foreach (char character in planId)
            {
                if (char.IsDigit(character))
                    return character.ToString();
            }

            return null;
        }

        private static List<string> LoadRoiIds(string roiListPath)
        {
            var roiIds = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string[] lines = File.ReadAllLines(roiListPath);

            for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
            {
                string raw = lines[lineIndex].Trim();

                if (raw.Length == 0 || raw.StartsWith("#"))
                    continue;

                List<string> cells = ParseCsvLine(raw, lineIndex + 1);
                if (cells.Count != 1)
                {
                    throw new InvalidDataException(
                        $"ROI list line {lineIndex + 1} must contain exactly one ROI ID.");
                }

                string roiId = cells[0].Trim();
                if (roiId.Equals("ROIID", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (roiId.Length == 0)
                    continue;

                if (seen.Add(roiId))
                    roiIds.Add(roiId);
            }

            return roiIds;
        }

        private static List<string> ParseCsvLine(string line, int lineNumber)
        {
            var cells = new List<string>();
            var value = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char character = line[i];

                if (character == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        value.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (character == ',' && !inQuotes)
                {
                    cells.Add(value.ToString());
                    value.Clear();
                }
                else
                {
                    value.Append(character);
                }
            }

            if (inQuotes)
                throw new InvalidDataException($"Unclosed quote on line {lineNumber}.");

            cells.Add(value.ToString());
            return cells;
        }

        private static PlanExtractionResult ExtractPlan(
            PatientPlanRequest request,
            PlanSetup plan,
            IList<string> roiIds)
        {
            var result = new PlanExtractionResult(request)
            {
                MatchedPlanId = plan.Id ?? ""
            };

            if (plan.StructureSet == null)
            {
                result.PlanStatus = "STRUCTURE_SET_MISSING";
                result.Message = "The plan has no structure set.";
                AddFailureRois(result, roiIds, result.PlanStatus);
                return result;
            }

            bool doseAvailable = plan.Dose != null && plan.IsDoseValid;
            var dvh = doseAvailable ? new DvhHelper(plan) : null;

            foreach (string requestedRoiId in roiIds)
            {
                var roiResult = new RoiExtractionResult(requestedRoiId);
                var structure = plan.StructureSet.Structures.FirstOrDefault(s =>
                    s != null && string.Equals(
                        s.Id,
                        requestedRoiId,
                        StringComparison.OrdinalIgnoreCase));

                if (structure == null)
                {
                    roiResult.Status = "ROI_NOT_FOUND";
                }
                else if (structure.IsEmpty)
                {
                    roiResult.Status = "ROI_EMPTY";
                }
                else
                {
                    roiResult.Volume = structure.Volume;

                    if (!IsFinite(roiResult.Volume.Value))
                    {
                        roiResult.Volume = null;
                        roiResult.Status = "ROI_INVALID_VOLUME";
                    }
                    else if (!doseAvailable)
                    {
                        roiResult.Status = "DOSE_UNAVAILABLE";
                    }
                    else
                    {
                        DvhMetricsResult metrics = dvh.GetMetrics(structure);
                        roiResult.D2 = metrics.D2Gy;
                        roiResult.D50 = metrics.D50Gy;
                        roiResult.D60 = metrics.D60Gy;
                        roiResult.Status = metrics.Status;
                    }
                }

                result.Rois.Add(roiResult);
            }

            var nonOkRois = result.Rois.Where(r => r.Status != "OK").ToList();

            if (!doseAvailable)
            {
                result.PlanStatus = "DOSE_UNAVAILABLE";
                result.Message = "The plan dose is missing or invalid; available ROI volumes were exported.";
            }
            else if (nonOkRois.Count > 0)
            {
                result.PlanStatus = "PARTIAL";
                result.Message = string.Join(
                    "; ",
                    nonOkRois.Select(r => r.RequestedRoiId + ":" + r.Status));
            }
            else
            {
                result.PlanStatus = "OK";
            }

            return result;
        }

        private static PlanExtractionResult CreateFailureResult(
            PatientPlanRequest request,
            IList<string> roiIds,
            string status,
            string message)
        {
            var result = new PlanExtractionResult(request)
            {
                PlanStatus = status,
                Message = message
            };

            AddFailureRois(result, roiIds, status);
            return result;
        }

        private static void AddFailureRois(
            PlanExtractionResult result,
            IEnumerable<string> roiIds,
            string status)
        {
            foreach (string roiId in roiIds)
            {
                result.Rois.Add(new RoiExtractionResult(roiId)
                {
                    Status = status
                });
            }
        }

        private static string BuildHeaderLine(IEnumerable<string> roiIds)
        {
            var columns = new List<string>
            {
                "PatientID",
                "CourseID",
                "PlanID",
                "MatchedPlanID",
                "PlanStatus",
                "Message",
                "VolumeUnit",
                "DoseUnit"
            };

            foreach (string roiId in roiIds)
            {
                columns.Add(roiId + "_Status");
                columns.Add(roiId + "_Vol");
                columns.Add(roiId + "_D2");
                columns.Add(roiId + "_D50");
                columns.Add(roiId + "_D60");
            }

            return CsvLine(columns);
        }

        private static string BuildOutputLine(PlanExtractionResult result)
        {
            var cells = new List<string>
            {
                result.PatientId,
                result.CourseId,
                result.PlanId,
                result.MatchedPlanId,
                result.PlanStatus,
                result.Message,
                "cm3",
                "Gy"
            };

            foreach (var roi in result.Rois)
            {
                cells.Add(roi.Status);
                cells.Add(NumberString(roi.Volume));
                cells.Add(NumberString(roi.D2));
                cells.Add(NumberString(roi.D50));
                cells.Add(NumberString(roi.D60));
            }

            return CsvLine(cells);
        }

        private static string CsvLine(IEnumerable<string> cells)
        {
            return string.Join(",", cells.Select(EscapeCsv));
        }

        private static void WriteOutputAtomically(string outputCsv, IEnumerable<string> rows)
        {
            string outputDirectory = Path.GetDirectoryName(outputCsv);
            string temporaryFile = Path.Combine(
                outputDirectory,
                "." + Path.GetFileName(outputCsv) + "." + Guid.NewGuid().ToString("N") + ".tmp");

            try
            {
                File.WriteAllLines(temporaryFile, rows, new UTF8Encoding(true));

                if (File.Exists(outputCsv))
                    File.Replace(temporaryFile, outputCsv, null);
                else
                    File.Move(temporaryFile, outputCsv);
            }
            finally
            {
                if (File.Exists(temporaryFile))
                    File.Delete(temporaryFile);
            }
        }

        private static string EscapeCsv(string value)
        {
            if (value == null)
                return "";

            bool mustQuote = value.Contains(",") ||
                             value.Contains("\"") ||
                             value.Contains("\n") ||
                             value.Contains("\r");
            string escaped = value.Replace("\"", "\"\"");
            return mustQuote ? "\"" + escaped + "\"" : escaped;
        }

        private static string NumberString(double? value)
        {
            return value.HasValue && IsFinite(value.Value)
                ? value.Value.ToString("0.###", CultureInfo.InvariantCulture)
                : "";
        }

        private static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static string AppendMessage(string current, string addition)
        {
            if (string.IsNullOrWhiteSpace(current))
                return addition ?? "";

            if (string.IsNullOrWhiteSpace(addition))
                return current;

            return current + " " + addition;
        }

        private sealed class PatientPlanRequest
        {
            public PatientPlanRequest(string patientId, string planId, string courseId)
            {
                PatientId = patientId;
                PlanId = planId;
                CourseId = courseId;
            }

            public string PatientId { get; }
            public string PlanId { get; }
            public string CourseId { get; }
        }

        private sealed class PlanExtractionResult
        {
            public PlanExtractionResult(PatientPlanRequest request)
            {
                PatientId = request.PatientId;
                CourseId = request.CourseId;
                PlanId = request.PlanId;
                Rois = new List<RoiExtractionResult>();
                MatchedPlanId = "";
                PlanStatus = "";
                Message = "";
            }

            public string PatientId { get; }
            public string CourseId { get; }
            public string PlanId { get; }
            public string MatchedPlanId { get; set; }
            public string PlanStatus { get; set; }
            public string Message { get; set; }
            public List<RoiExtractionResult> Rois { get; }
        }

        private sealed class RoiExtractionResult
        {
            public RoiExtractionResult(string requestedRoiId)
            {
                RequestedRoiId = requestedRoiId;
                Status = "";
            }

            public string RequestedRoiId { get; }
            public string Status { get; set; }
            public double? Volume { get; set; }
            public double? D2 { get; set; }
            public double? D50 { get; set; }
            public double? D60 { get; set; }
        }

        private sealed class DvhMetricsResult
        {
            public DvhMetricsResult(string status)
            {
                Status = status;
            }

            public string Status { get; set; }
            public double? D2Gy { get; set; }
            public double? D50Gy { get; set; }
            public double? D60Gy { get; set; }
        }

        private sealed class DvhHelper
        {
            private readonly PlanSetup _plan;

            public DvhHelper(PlanSetup plan)
            {
                _plan = plan;
            }

            public DvhMetricsResult GetMetrics(Structure structure)
            {
                try
                {
                    if (structure == null || structure.IsEmpty)
                        return new DvhMetricsResult("DVH_UNAVAILABLE");

                    double? binWidth = GetBinWidthForPlanDoseUnit();
                    if (!binWidth.HasValue)
                        return new DvhMetricsResult("DOSE_UNIT_UNSUPPORTED");

                    var dvh = _plan.GetDVHCumulativeData(
                        structure,
                        DoseValuePresentation.Absolute,
                        VolumePresentation.Relative,
                        binWidth.Value);

                    if (dvh == null || dvh.CurveData == null || !dvh.CurveData.Any())
                        return new DvhMetricsResult("DVH_UNAVAILABLE");

                    if (dvh.CurveData.Any(point =>
                        point.DoseValue.Unit != DoseValue.DoseUnit.Gy &&
                        point.DoseValue.Unit != DoseValue.DoseUnit.cGy))
                    {
                        return new DvhMetricsResult("DOSE_UNIT_UNSUPPORTED");
                    }

                    if (!IsFinite(dvh.Coverage) ||
                        !IsFinite(dvh.SamplingCoverage) ||
                        dvh.Coverage < MinimumDvhCoverage ||
                        dvh.SamplingCoverage < MinimumDvhCoverage)
                    {
                        return new DvhMetricsResult("DVH_INCOMPLETE_COVERAGE");
                    }

                    var result = new DvhMetricsResult("")
                    {
                        D2Gy = DoseAtVolumePercentGy(dvh.CurveData, 2),
                        D50Gy = DoseAtVolumePercentGy(dvh.CurveData, 50),
                        D60Gy = DoseAtVolumePercentGy(dvh.CurveData, 60)
                    };

                    if (result.D2Gy.HasValue &&
                        result.D50Gy.HasValue &&
                        result.D60Gy.HasValue)
                    {
                        result.Status = "OK";
                    }
                    else if (result.D2Gy.HasValue ||
                             result.D50Gy.HasValue ||
                             result.D60Gy.HasValue)
                    {
                        result.Status = "DVH_PARTIAL";
                    }
                    else
                    {
                        result.Status = "DVH_UNAVAILABLE";
                    }

                    return result;
                }
                catch
                {
                    return new DvhMetricsResult("DVH_UNAVAILABLE");
                }
            }

            private double? GetBinWidthForPlanDoseUnit()
            {
                DoseValue.DoseUnit unit = _plan.Dose.DoseMax3D.Unit;

                if (unit == DoseValue.DoseUnit.Gy)
                    return DvhBinWidthGy;

                if (unit == DoseValue.DoseUnit.cGy)
                    return DvhBinWidthGy * 100.0;

                return null;
            }

            private static double? DoseAtVolumePercentGy(
                IEnumerable<DVHPoint> curve,
                double volumePercent)
            {
                DVHPoint? previous = null;

                foreach (var point in curve)
                {
                    if (!IsFinite(point.Volume))
                        return null;

                    if (point.Volume <= volumePercent)
                    {
                        if (!previous.HasValue)
                        {
                            return Math.Abs(point.Volume - volumePercent) < 1e-6
                                ? DoseToGy(point.DoseValue)
                                : null;
                        }

                        double firstVolume = previous.Value.Volume;
                        double secondVolume = point.Volume;

                        if (firstVolume < volumePercent || secondVolume > volumePercent)
                            return null;

                        double? firstDoseGy = DoseToGy(previous.Value.DoseValue);
                        double? secondDoseGy = DoseToGy(point.DoseValue);

                        if (!firstDoseGy.HasValue || !secondDoseGy.HasValue)
                            return null;

                        if (Math.Abs(secondVolume - firstVolume) < 1e-6)
                            return secondDoseGy.Value;

                        double fraction =
                            (volumePercent - firstVolume) / (secondVolume - firstVolume);
                        double interpolated =
                            firstDoseGy.Value + fraction * (secondDoseGy.Value - firstDoseGy.Value);

                        return IsFinite(interpolated) ? (double?)interpolated : null;
                    }

                    previous = point;
                }

                return null;
            }

            private static double? DoseToGy(DoseValue doseValue)
            {
                if (doseValue.IsUndefined() || !IsFinite(doseValue.Dose))
                    return null;

                if (doseValue.Unit == DoseValue.DoseUnit.Gy)
                    return doseValue.Dose;

                if (doseValue.Unit == DoseValue.DoseUnit.cGy)
                    return doseValue.Dose / 100.0;

                return null;
            }
        }
    }
}
