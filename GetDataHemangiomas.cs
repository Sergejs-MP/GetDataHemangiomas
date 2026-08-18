using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: AssemblyVersion("1.0.0.2")]
[assembly: AssemblyFileVersion("1.0.0.2")]
[assembly: AssemblyInformationalVersion("1.2.0")]

namespace GetDataHemangiomas
{
    class Program
    {
        private const string ExtractorVersion = "1.2.0";
        private const string DoseBasis = "PHYSICAL_ABSOLUTE";
        private const string FallbackMethod = "PHYSICAL_LINE_HISTOGRAM";
        private const double DvhBinWidthGy = 0.01;
        private const double MinimumDvhCoverage = 0.999;
        private const int LineHistogramBinCount = 1024;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("GetDataHemangiomas " + ExtractorVersion);

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
            var lineSampler = doseAvailable && dvh.DosePresentationStateSafe
                ? new LineDoseSampler(plan)
                : null;

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
                    try
                    {
                        double volume = structure.Volume;
                        if (IsFinite(volume) && volume > 0.0)
                            roiResult.Volume = volume;
                        else
                            roiResult.AddWarning("ROI_INVALID_VOLUME");
                    }
                    catch
                    {
                        roiResult.AddWarning("ROI_VOLUME_UNAVAILABLE");
                    }

                    if (!doseAvailable)
                    {
                        roiResult.Status = "DOSE_UNAVAILABLE";
                    }
                    else
                    {
                        DvhMetricsResult metrics = dvh.GetMetrics(structure);
                        roiResult.DvhStatus = metrics.Status;
                        roiResult.DvhCoverage = metrics.Coverage;
                        roiResult.DvhSamplingCoverage = metrics.SamplingCoverage;
                        roiResult.AddWarnings(metrics.WarningCodes);

                        if (metrics.HasCompleteDoseMetrics)
                        {
                            roiResult.D2 = metrics.D2Gy;
                            roiResult.D50 = metrics.D50Gy;
                            roiResult.D60 = metrics.D60Gy;
                            roiResult.DoseSource = "DVH";
                            roiResult.LineStatus = "NOT_NEEDED";
                        }
                        else
                        {
                            LineDoseMetricsResult line;
                            if (lineSampler == null)
                            {
                                line = new LineDoseMetricsResult(
                                    "PRESENTATION_RESTORE_ERROR");
                                line.AddWarning(
                                    "LINE_BLOCKED_BY_DVH_PRESENTATION_ERROR");
                            }
                            else
                            {
                                line = lineSampler.GetMetrics(structure);
                            }
                            roiResult.LineStatus = line.Status;
                            roiResult.LineInsideSamples = line.InsideSamples;
                            roiResult.LineValidDoseSamples = line.ValidDoseSamples;
                            roiResult.LineSamplingCoverage = line.SamplingCoverage;
                            roiResult.LineInsideVolumeEstimate =
                                line.InsideVolumeEstimateCc;
                            roiResult.LineVolumeRatio = line.SampledVolumeRatio;
                            roiResult.LineMaxDose = line.MaxDoseGy;
                            roiResult.LineBinWidth = line.BinWidthGy;
                            roiResult.AddWarnings(line.WarningCodes);

                            if (line.HasCompleteDoseMetrics)
                            {
                                roiResult.D2 = line.D2Gy;
                                roiResult.D50 = line.D50Gy;
                                roiResult.D60 = line.D60Gy;
                                roiResult.DoseSource = "LINE";
                                roiResult.AddWarning("LINE_FALLBACK");
                            }
                            else
                            {
                                roiResult.D2 = metrics.D2Gy;
                                roiResult.D50 = metrics.D50Gy;
                                roiResult.D60 = metrics.D60Gy;
                                roiResult.DoseSource = metrics.HasAnyDoseMetric
                                    ? "DVH_PARTIAL"
                                    : "NONE";
                                roiResult.AddWarning("LINE_FALLBACK_UNAVAILABLE");
                            }
                        }

                        if (!roiResult.HasAnyDoseMetric)
                            roiResult.Status = "DOSE_UNAVAILABLE";
                        else if (!roiResult.HasAllRequestedValues)
                            roiResult.Status = "PARTIAL";
                        else if (roiResult.WarningCodes.Count > 0)
                            roiResult.Status = "WARNING";
                        else
                            roiResult.Status = "OK";
                    }
                }

                result.Rois.Add(roiResult);
            }

            var incompleteRois = result.Rois.Where(
                r => !r.HasAllRequestedValues).ToList();
            var warningRois = result.Rois.Where(
                r => r.HasAllRequestedValues && r.WarningCodes.Count > 0).ToList();

            if (!doseAvailable)
            {
                result.PlanStatus = "DOSE_UNAVAILABLE";
                result.Message = "The plan dose is missing or invalid; available ROI volumes were exported.";
            }
            else if (incompleteRois.Count > 0)
            {
                result.PlanStatus = "PARTIAL";
                result.Message = string.Join(
                    "; ",
                    result.Rois
                        .Where(r => r.Status != "OK")
                        .Select(r => r.RequestedRoiId + ":" + r.Status));
            }
            else if (warningRois.Count > 0)
            {
                result.PlanStatus = "WARNING";
                result.Message = string.Join(
                    "; ",
                    warningRois.Select(
                        r => r.RequestedRoiId + ":" + r.Status));
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
                "ExtractorVersion",
                "DoseBasis",
                "FallbackMethod",
                "VolumeUnit",
                "DoseUnit"
            };

            foreach (string roiId in roiIds)
            {
                columns.Add(roiId + "_Status");
                columns.Add(roiId + "_WarningCodes");
                columns.Add(roiId + "_DoseSource");
                columns.Add(roiId + "_Vol");
                columns.Add(roiId + "_D2");
                columns.Add(roiId + "_D50");
                columns.Add(roiId + "_D60");
                columns.Add(roiId + "_DVHStatus");
                columns.Add(roiId + "_DVHCoverage");
                columns.Add(roiId + "_DVHSamplingCoverage");
                columns.Add(roiId + "_LineStatus");
                columns.Add(roiId + "_LineInsideSamples");
                columns.Add(roiId + "_LineValidDoseSamples");
                columns.Add(roiId + "_LineSamplingCoverage");
                columns.Add(roiId + "_LineInsideVolumeEstimate");
                columns.Add(roiId + "_LineVolumeRatio");
                columns.Add(roiId + "_LineMaxDose");
                columns.Add(roiId + "_LineBinWidth");
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
                ExtractorVersion,
                DoseBasis,
                FallbackMethod,
                "cm3",
                "Gy"
            };

            foreach (var roi in result.Rois)
            {
                cells.Add(roi.Status);
                cells.Add(string.Join("|", roi.WarningCodes));
                cells.Add(roi.DoseSource);
                cells.Add(NumberString(roi.Volume));
                cells.Add(NumberString(roi.D2));
                cells.Add(NumberString(roi.D50));
                cells.Add(NumberString(roi.D60));
                cells.Add(roi.DvhStatus);
                cells.Add(CoverageString(roi.DvhCoverage));
                cells.Add(CoverageString(roi.DvhSamplingCoverage));
                cells.Add(roi.LineStatus);
                cells.Add(LongString(roi.LineInsideSamples));
                cells.Add(LongString(roi.LineValidDoseSamples));
                cells.Add(CoverageString(roi.LineSamplingCoverage));
                cells.Add(AuditNumberString(roi.LineInsideVolumeEstimate));
                cells.Add(CoverageString(roi.LineVolumeRatio));
                cells.Add(AuditNumberString(roi.LineMaxDose));
                cells.Add(AuditNumberString(roi.LineBinWidth));
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

        private static string CoverageString(double? value)
        {
            return value.HasValue && IsFinite(value.Value)
                ? value.Value.ToString("R", CultureInfo.InvariantCulture)
                : "";
        }

        private static string AuditNumberString(double? value)
        {
            return value.HasValue && IsFinite(value.Value)
                ? value.Value.ToString("R", CultureInfo.InvariantCulture)
                : "";
        }

        private static string LongString(long? value)
        {
            return value.HasValue
                ? value.Value.ToString(CultureInfo.InvariantCulture)
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
                DoseSource = "";
                DvhStatus = "";
                LineStatus = "";
                WarningCodes = new List<string>();
            }

            public string RequestedRoiId { get; }
            public string Status { get; set; }
            public string DoseSource { get; set; }
            public string DvhStatus { get; set; }
            public string LineStatus { get; set; }
            public List<string> WarningCodes { get; }
            public double? Volume { get; set; }
            public double? DvhCoverage { get; set; }
            public double? DvhSamplingCoverage { get; set; }
            public double? D2 { get; set; }
            public double? D50 { get; set; }
            public double? D60 { get; set; }
            public long? LineInsideSamples { get; set; }
            public long? LineValidDoseSamples { get; set; }
            public double? LineSamplingCoverage { get; set; }
            public double? LineInsideVolumeEstimate { get; set; }
            public double? LineVolumeRatio { get; set; }
            public double? LineMaxDose { get; set; }
            public double? LineBinWidth { get; set; }

            public bool HasAnyDoseMetric
            {
                get
                {
                    return D2.HasValue || D50.HasValue || D60.HasValue;
                }
            }

            public bool HasCompleteDoseMetrics
            {
                get
                {
                    return D2.HasValue && D50.HasValue && D60.HasValue;
                }
            }

            public bool HasAllRequestedValues
            {
                get
                {
                    return Volume.HasValue && HasCompleteDoseMetrics;
                }
            }

            public void AddWarning(string warningCode)
            {
                if (!string.IsNullOrWhiteSpace(warningCode) &&
                    !WarningCodes.Contains(warningCode))
                {
                    WarningCodes.Add(warningCode);
                }
            }

            public void AddWarnings(IEnumerable<string> warningCodes)
            {
                if (warningCodes == null)
                    return;

                foreach (string warningCode in warningCodes)
                    AddWarning(warningCode);
            }
        }

        private sealed class DvhMetricsResult
        {
            public DvhMetricsResult(string status)
            {
                Status = status;
                WarningCodes = new List<string>();
            }

            public string Status { get; set; }
            public List<string> WarningCodes { get; }
            public double? Coverage { get; set; }
            public double? SamplingCoverage { get; set; }
            public double? D2Gy { get; set; }
            public double? D50Gy { get; set; }
            public double? D60Gy { get; set; }

            public bool HasAnyDoseMetric
            {
                get
                {
                    return D2Gy.HasValue || D50Gy.HasValue || D60Gy.HasValue;
                }
            }

            public bool HasCompleteDoseMetrics
            {
                get
                {
                    return D2Gy.HasValue && D50Gy.HasValue && D60Gy.HasValue;
                }
            }

            public void AddWarning(string warningCode)
            {
                if (!string.IsNullOrWhiteSpace(warningCode) &&
                    !WarningCodes.Contains(warningCode))
                {
                    WarningCodes.Add(warningCode);
                }
            }
        }

        private sealed class LineDoseMetricsResult
        {
            public LineDoseMetricsResult(string status)
            {
                Status = status;
                WarningCodes = new List<string>();
            }

            public string Status { get; set; }
            public List<string> WarningCodes { get; }
            public double? D2Gy { get; set; }
            public double? D50Gy { get; set; }
            public double? D60Gy { get; set; }
            public long? InsideSamples { get; set; }
            public long? ValidDoseSamples { get; set; }
            public double? SamplingCoverage { get; set; }
            public double? InsideVolumeEstimateCc { get; set; }
            public double? SampledVolumeRatio { get; set; }
            public double? MaxDoseGy { get; set; }
            public double? BinWidthGy { get; set; }

            public bool HasCompleteDoseMetrics
            {
                get
                {
                    return D2Gy.HasValue && D50Gy.HasValue && D60Gy.HasValue;
                }
            }

            public void AddWarning(string warningCode)
            {
                if (!string.IsNullOrWhiteSpace(warningCode) &&
                    !WarningCodes.Contains(warningCode))
                {
                    WarningCodes.Add(warningCode);
                }
            }

            public void AddWarnings(IEnumerable<string> warningCodes)
            {
                if (warningCodes == null)
                    return;

                foreach (string warningCode in warningCodes)
                    AddWarning(warningCode);
            }

            public void ClearDoseMetrics()
            {
                D2Gy = null;
                D50Gy = null;
                D60Gy = null;
            }
        }

        private sealed class LineSamplingGeometry
        {
            public int XCount { get; set; }
            public int YCount { get; set; }
            public int ZCount { get; set; }
            public double XStart { get; set; }
            public double YStart { get; set; }
            public double ZStart { get; set; }
            public double XSpacing { get; set; }
            public double YSpacing { get; set; }
            public double ZSpacing { get; set; }
            public double ZEnd { get; set; }

            public double CellVolumeCc
            {
                get
                {
                    return XSpacing * YSpacing * ZSpacing / 1000.0;
                }
            }
        }

        private sealed class LineScanResult
        {
            public LineScanResult()
            {
                Completed = true;
                SampleFingerprint = 14695981039346656037UL;
                WarningCodes = new List<string>();
            }

            public bool Completed { get; set; }
            public bool SawUnsupportedDoseUnit { get; set; }
            public long InsideSamples { get; set; }
            public long ValidDoseSamples { get; set; }
            public double MaxDoseGy { get; set; }
            public ulong SampleFingerprint { get; set; }
            public List<string> WarningCodes { get; }

            public void AddWarning(string warningCode)
            {
                if (!string.IsNullOrWhiteSpace(warningCode) &&
                    !WarningCodes.Contains(warningCode))
                {
                    WarningCodes.Add(warningCode);
                }
            }
        }

        private sealed class LineDoseSampler
        {
            private readonly PlanSetup _plan;
            private bool _unusable;

            public LineDoseSampler(PlanSetup plan)
            {
                _plan = plan;
            }

            public LineDoseMetricsResult GetMetrics(Structure structure)
            {
                if (_unusable)
                {
                    var unavailable =
                        new LineDoseMetricsResult("PRESENTATION_RESTORE_ERROR");
                    unavailable.AddWarning(
                        "LINE_DOSE_PRESENTATION_RESTORE_ERROR");
                    return unavailable;
                }

                if (structure == null || structure.IsEmpty || !structure.HasSegment)
                {
                    var unavailable = new LineDoseMetricsResult("STRUCTURE_UNAVAILABLE");
                    unavailable.AddWarning("LINE_STRUCTURE_UNAVAILABLE");
                    return unavailable;
                }

                if (_plan.Dose == null || !_plan.IsDoseValid)
                {
                    var unavailable = new LineDoseMetricsResult("DOSE_UNAVAILABLE");
                    unavailable.AddWarning("LINE_DOSE_UNAVAILABLE");
                    return unavailable;
                }

                LineSamplingGeometry geometry;
                string geometryWarning;
                try
                {
                    if (!TryCreateGeometry(
                        structure,
                        out geometry,
                        out geometryWarning))
                    {
                        var invalidGeometry =
                            new LineDoseMetricsResult("GEOMETRY_INVALID");
                        invalidGeometry.AddWarning(geometryWarning);
                        return invalidGeometry;
                    }
                }
                catch
                {
                    var invalidGeometry = new LineDoseMetricsResult("GEOMETRY_INVALID");
                    invalidGeometry.AddWarning("LINE_GEOMETRY_ERROR");
                    return invalidGeometry;
                }

                DoseValuePresentation originalPresentation;
                try
                {
                    originalPresentation = _plan.DoseValuePresentation;
                }
                catch
                {
                    var presentationUnavailable =
                        new LineDoseMetricsResult("PRESENTATION_UNAVAILABLE");
                    presentationUnavailable.AddWarning(
                        "LINE_DOSE_PRESENTATION_UNAVAILABLE");
                    return presentationUnavailable;
                }

                bool presentationChanged =
                    originalPresentation != DoseValuePresentation.Absolute;
                LineDoseMetricsResult result = null;
                bool restoreFailed = false;

                try
                {
                    if (presentationChanged)
                        _plan.DoseValuePresentation = DoseValuePresentation.Absolute;

                    result = SampleAbsoluteDose(structure, geometry);
                }
                catch
                {
                    result = new LineDoseMetricsResult("ERROR");
                    result.AddWarning("LINE_EXTRACTION_ERROR");
                }
                finally
                {
                    if (presentationChanged)
                    {
                        try
                        {
                            _plan.DoseValuePresentation = originalPresentation;
                        }
                        catch
                        {
                            restoreFailed = true;
                        }
                    }
                }

                if (result == null)
                {
                    result = new LineDoseMetricsResult("ERROR");
                    result.AddWarning("LINE_EXTRACTION_ERROR");
                }

                if (restoreFailed)
                {
                    _unusable = true;
                    result.ClearDoseMetrics();
                    result.Status = "PRESENTATION_RESTORE_ERROR";
                    result.AddWarning("LINE_DOSE_PRESENTATION_RESTORE_ERROR");
                }

                return result;
            }

            private LineDoseMetricsResult SampleAbsoluteDose(
                Structure structure,
                LineSamplingGeometry geometry)
            {
                var result = new LineDoseMetricsResult("");
                LineScanResult firstPass = Scan(
                    structure,
                    geometry,
                    null,
                    0.0,
                    0.0);

                result.AddWarnings(firstPass.WarningCodes);

                if (!firstPass.Completed)
                {
                    result.Status = "STRUCTURE_PROFILE_UNAVAILABLE";
                    return result;
                }

                result.InsideSamples = firstPass.InsideSamples;
                result.ValidDoseSamples = firstPass.ValidDoseSamples;

                if (firstPass.InsideSamples > 0)
                {
                    result.SamplingCoverage =
                        (double)firstPass.ValidDoseSamples /
                        firstPass.InsideSamples;

                    double sampledVolume =
                        firstPass.InsideSamples * geometry.CellVolumeCc;
                    if (IsFinite(sampledVolume) && sampledVolume >= 0.0)
                        result.InsideVolumeEstimateCc = sampledVolume;

                    try
                    {
                        double structureVolume = structure.Volume;
                        if (result.InsideVolumeEstimateCc.HasValue &&
                            IsFinite(structureVolume) &&
                            structureVolume > 0.0)
                        {
                            result.SampledVolumeRatio =
                                result.InsideVolumeEstimateCc.Value /
                                structureVolume;
                        }
                    }
                    catch
                    {
                        result.AddWarning("LINE_VOLUME_RATIO_UNAVAILABLE");
                    }
                }

                if (firstPass.InsideSamples == 0)
                {
                    result.Status = "NO_INSIDE_SAMPLES";
                    result.AddWarning("LINE_NO_INSIDE_SAMPLES");
                    return result;
                }

                if (firstPass.SawUnsupportedDoseUnit)
                {
                    result.Status = "DOSE_UNIT_UNSUPPORTED";
                    return result;
                }

                if (firstPass.ValidDoseSamples == 0)
                {
                    result.Status = "NO_VALID_DOSE";
                    result.AddWarning("LINE_NO_VALID_DOSE");
                    return result;
                }

                if (firstPass.ValidDoseSamples < firstPass.InsideSamples)
                    result.AddWarning("LINE_SAMPLING_INCOMPLETE");

                result.MaxDoseGy = firstPass.MaxDoseGy;

                if (firstPass.MaxDoseGy == 0.0)
                {
                    result.BinWidthGy = 0.0;
                    result.D2Gy = 0.0;
                    result.D50Gy = 0.0;
                    result.D60Gy = 0.0;
                    result.Status = result.WarningCodes.Count == 0
                        ? "OK"
                        : "WARNING";
                    return result;
                }

                double binWidth = firstPass.MaxDoseGy / LineHistogramBinCount;
                if (!IsFinite(binWidth) || binWidth <= 0.0)
                {
                    result.Status = "HISTOGRAM_INVALID";
                    result.AddWarning("LINE_HISTOGRAM_BIN_WIDTH_INVALID");
                    return result;
                }

                result.BinWidthGy = binWidth;
                var histogram = new long[LineHistogramBinCount];
                LineScanResult secondPass = Scan(
                    structure,
                    geometry,
                    histogram,
                    binWidth,
                    firstPass.MaxDoseGy);
                result.AddWarnings(secondPass.WarningCodes);

                if (secondPass.SawUnsupportedDoseUnit)
                {
                    result.Status = "DOSE_UNIT_UNSUPPORTED";
                    return result;
                }

                if (!secondPass.Completed ||
                    secondPass.InsideSamples != firstPass.InsideSamples ||
                    secondPass.ValidDoseSamples != firstPass.ValidDoseSamples ||
                    secondPass.SampleFingerprint != firstPass.SampleFingerprint ||
                    Math.Abs(secondPass.MaxDoseGy - firstPass.MaxDoseGy) >
                        Math.Max(1e-9, Math.Abs(firstPass.MaxDoseGy) * 1e-9))
                {
                    result.Status = "PASS_MISMATCH";
                    result.AddWarning("LINE_SAMPLING_PASS_MISMATCH");
                    return result;
                }

                long histogramSamples = 0;
                foreach (long count in histogram)
                {
                    if (long.MaxValue - histogramSamples < count)
                    {
                        result.Status = "HISTOGRAM_INVALID";
                        result.AddWarning("LINE_HISTOGRAM_COUNT_OVERFLOW");
                        return result;
                    }

                    histogramSamples += count;
                }

                if (histogramSamples != firstPass.ValidDoseSamples)
                {
                    result.Status = "PASS_MISMATCH";
                    result.AddWarning("LINE_HISTOGRAM_SAMPLE_MISMATCH");
                    return result;
                }

                result.D2Gy = DoseAtVolumeFromHistogram(
                    histogram,
                    histogramSamples,
                    binWidth,
                    firstPass.MaxDoseGy,
                    2.0);
                result.D50Gy = DoseAtVolumeFromHistogram(
                    histogram,
                    histogramSamples,
                    binWidth,
                    firstPass.MaxDoseGy,
                    50.0);
                result.D60Gy = DoseAtVolumeFromHistogram(
                    histogram,
                    histogramSamples,
                    binWidth,
                    firstPass.MaxDoseGy,
                    60.0);

                if (!result.HasCompleteDoseMetrics)
                {
                    result.Status = "HISTOGRAM_INVALID";
                    result.AddWarning("LINE_HISTOGRAM_METRICS_UNAVAILABLE");
                    return result;
                }

                result.Status = result.WarningCodes.Count == 0
                    ? "OK"
                    : "WARNING";
                return result;
            }

            private LineScanResult Scan(
                Structure structure,
                LineSamplingGeometry geometry,
                long[] histogram,
                double binWidth,
                double histogramMaxDoseGy)
            {
                var scan = new LineScanResult();

                for (int ix = 0; ix < geometry.XCount; ix++)
                {
                    double x = geometry.XStart + ix * geometry.XSpacing;

                    for (int iy = 0; iy < geometry.YCount; iy++)
                    {
                        double y = geometry.YStart + iy * geometry.YSpacing;
                        var start = new VVector(x, y, geometry.ZStart);
                        var end = new VVector(x, y, geometry.ZEnd);
                        SegmentProfile segment;

                        try
                        {
                            segment = structure.GetSegmentProfile(
                                start,
                                end,
                                new BitArray(geometry.ZCount));
                        }
                        catch
                        {
                            scan.Completed = false;
                            scan.AddWarning("LINE_SEGMENT_PROFILE_ERROR");
                            return scan;
                        }

                        if (segment == null || segment.Count == 0)
                        {
                            for (int sampleIndex = 0;
                                sampleIndex < geometry.ZCount;
                                sampleIndex++)
                            {
                                scan.SampleFingerprint = UpdateFingerprint(
                                    scan.SampleFingerprint,
                                    0UL);
                            }

                            continue;
                        }

                        if (segment.Count != geometry.ZCount)
                        {
                            scan.Completed = false;
                            scan.AddWarning("LINE_SEGMENT_PROFILE_INCOMPLETE");
                            return scan;
                        }

                        bool hasInsideSample = false;
                        for (int sampleIndex = 0;
                            sampleIndex < geometry.ZCount;
                            sampleIndex++)
                        {
                            if (segment[sampleIndex].Value)
                            {
                                hasInsideSample = true;
                                break;
                            }
                        }

                        if (!hasInsideSample)
                        {
                            for (int sampleIndex = 0;
                                sampleIndex < geometry.ZCount;
                                sampleIndex++)
                            {
                                scan.SampleFingerprint = UpdateFingerprint(
                                    scan.SampleFingerprint,
                                    0UL);
                            }

                            continue;
                        }

                        DoseProfile doseProfile = null;
                        bool doseProfileFailed = false;

                        try
                        {
                            doseProfile = _plan.Dose.GetDoseProfile(
                                start,
                                end,
                                new double[geometry.ZCount]);
                        }
                        catch
                        {
                            doseProfileFailed = true;
                            scan.AddWarning("LINE_DOSE_PROFILE_ERROR");
                        }

                        if (doseProfile == null)
                        {
                            doseProfileFailed = true;
                            scan.AddWarning("LINE_DOSE_PROFILE_UNAVAILABLE");
                        }
                        else if (doseProfile.Count != geometry.ZCount)
                        {
                            doseProfileFailed = true;
                            scan.AddWarning("LINE_DOSE_PROFILE_INCOMPLETE");
                        }

                        bool supportedUnit =
                            doseProfile != null &&
                            (doseProfile.Unit == DoseValue.DoseUnit.Gy ||
                             doseProfile.Unit == DoseValue.DoseUnit.cGy);

                        if (doseProfile != null && !supportedUnit)
                        {
                            doseProfileFailed = true;
                            scan.SawUnsupportedDoseUnit = true;
                            scan.AddWarning("LINE_DOSE_UNIT_UNSUPPORTED");
                        }

                        for (int sampleIndex = 0;
                            sampleIndex < geometry.ZCount;
                            sampleIndex++)
                        {
                            if (!segment[sampleIndex].Value)
                            {
                                scan.SampleFingerprint = UpdateFingerprint(
                                    scan.SampleFingerprint,
                                    0UL);
                                continue;
                            }

                            if (scan.InsideSamples == long.MaxValue)
                            {
                                scan.Completed = false;
                                scan.AddWarning("LINE_SAMPLE_COUNT_OVERFLOW");
                                return scan;
                            }

                            scan.InsideSamples++;

                            if (doseProfileFailed ||
                                doseProfile == null ||
                                sampleIndex >= doseProfile.Count)
                            {
                                scan.SampleFingerprint = UpdateFingerprint(
                                    scan.SampleFingerprint,
                                    1UL);
                                continue;
                            }

                            double doseGy = doseProfile[sampleIndex].Value;
                            if (doseProfile.Unit == DoseValue.DoseUnit.cGy)
                                doseGy /= 100.0;

                            if (!IsFinite(doseGy) || doseGy < 0.0)
                            {
                                scan.AddWarning("LINE_DOSE_SAMPLES_REJECTED");
                                scan.SampleFingerprint = UpdateFingerprint(
                                    scan.SampleFingerprint,
                                    1UL);
                                continue;
                            }

                            scan.SampleFingerprint = UpdateFingerprint(
                                scan.SampleFingerprint,
                                2UL);
                            scan.SampleFingerprint = UpdateFingerprint(
                                scan.SampleFingerprint,
                                unchecked((ulong)BitConverter.DoubleToInt64Bits(doseGy)));

                            if (scan.ValidDoseSamples == long.MaxValue)
                            {
                                scan.Completed = false;
                                scan.AddWarning("LINE_SAMPLE_COUNT_OVERFLOW");
                                return scan;
                            }

                            scan.ValidDoseSamples++;
                            if (doseGy > scan.MaxDoseGy)
                                scan.MaxDoseGy = doseGy;

                            if (histogram == null)
                                continue;

                            double tolerance = Math.Max(
                                1e-9,
                                Math.Abs(histogramMaxDoseGy) * 1e-9);
                            if (doseGy > histogramMaxDoseGy + tolerance)
                            {
                                scan.Completed = false;
                                scan.AddWarning("LINE_SAMPLING_PASS_MAX_CHANGED");
                                return scan;
                            }

                            int bin = (int)Math.Floor(doseGy / binWidth);
                            if (bin < 0)
                                bin = 0;
                            if (bin >= histogram.Length)
                                bin = histogram.Length - 1;

                            if (histogram[bin] == long.MaxValue)
                            {
                                scan.Completed = false;
                                scan.AddWarning("LINE_HISTOGRAM_COUNT_OVERFLOW");
                                return scan;
                            }

                            histogram[bin]++;
                        }
                    }
                }

                return scan;
            }

            private static ulong UpdateFingerprint(ulong current, ulong value)
            {
                unchecked
                {
                    current ^= value;
                    current *= 1099511628211UL;
                    return current;
                }
            }

            private bool TryCreateGeometry(
                Structure structure,
                out LineSamplingGeometry geometry,
                out string warningCode)
            {
                geometry = null;
                warningCode = null;

                double xResolution = _plan.Dose.XRes;
                double yResolution = _plan.Dose.YRes;
                double zResolution = _plan.Dose.ZRes;

                if (!IsFinitePositive(xResolution) ||
                    !IsFinitePositive(yResolution) ||
                    !IsFinitePositive(zResolution))
                {
                    warningCode = "LINE_DOSE_GRID_INVALID";
                    return false;
                }

                var bounds = structure.MeshGeometry.Bounds;
                if (!IsFinite(bounds.X) ||
                    !IsFinite(bounds.Y) ||
                    !IsFinite(bounds.Z) ||
                    !IsFinitePositive(bounds.SizeX) ||
                    !IsFinitePositive(bounds.SizeY) ||
                    !IsFinitePositive(bounds.SizeZ))
                {
                    warningCode = "LINE_STRUCTURE_BOUNDS_INVALID";
                    return false;
                }

                int xCount;
                int yCount;
                int zCount;
                if (!TryGetSampleCount(bounds.SizeX, xResolution, 1, out xCount) ||
                    !TryGetSampleCount(bounds.SizeY, yResolution, 1, out yCount) ||
                    !TryGetSampleCount(bounds.SizeZ, zResolution, 2, out zCount))
                {
                    warningCode = "LINE_SAMPLE_GRID_TOO_LARGE";
                    return false;
                }

                long lineCount = (long)xCount * yCount;
                if (lineCount <= 0 ||
                    lineCount > long.MaxValue / zCount)
                {
                    warningCode = "LINE_SAMPLE_GRID_TOO_LARGE";
                    return false;
                }

                double xSpacing = bounds.SizeX / xCount;
                double ySpacing = bounds.SizeY / yCount;
                double zSpacing = bounds.SizeZ / zCount;
                if (!IsFinitePositive(xSpacing) ||
                    !IsFinitePositive(ySpacing) ||
                    !IsFinitePositive(zSpacing))
                {
                    warningCode = "LINE_SAMPLE_GRID_INVALID";
                    return false;
                }

                geometry = new LineSamplingGeometry
                {
                    XCount = xCount,
                    YCount = yCount,
                    ZCount = zCount,
                    XStart = bounds.X + 0.5 * xSpacing,
                    YStart = bounds.Y + 0.5 * ySpacing,
                    ZStart = bounds.Z + 0.5 * zSpacing,
                    XSpacing = xSpacing,
                    YSpacing = ySpacing,
                    ZSpacing = zSpacing,
                    ZEnd = bounds.Z + bounds.SizeZ - 0.5 * zSpacing
                };

                bool validGeometry =
                    IsFinite(geometry.XStart) &&
                    IsFinite(geometry.YStart) &&
                    IsFinite(geometry.ZStart) &&
                    IsFinite(geometry.ZEnd) &&
                    IsFinitePositive(geometry.CellVolumeCc);

                if (!validGeometry)
                    warningCode = "LINE_SAMPLE_GRID_INVALID";

                return validGeometry;
            }

            private static bool TryGetSampleCount(
                double size,
                double resolution,
                int minimum,
                out int sampleCount)
            {
                sampleCount = 0;
                double rawCount = Math.Ceiling(size / resolution);

                if (!IsFinite(rawCount) || rawCount > int.MaxValue)
                    return false;

                sampleCount = Math.Max(minimum, (int)rawCount);
                return sampleCount > 0;
            }

            private static double? DoseAtVolumeFromHistogram(
                long[] histogram,
                long totalSamples,
                double binWidthGy,
                double maxDoseGy,
                double volumePercent)
            {
                if (histogram == null ||
                    histogram.Length == 0 ||
                    totalSamples <= 0 ||
                    !IsFinite(binWidthGy) ||
                    binWidthGy <= 0.0 ||
                    !IsFinite(maxDoseGy) ||
                    maxDoseGy < 0.0 ||
                    !IsFinite(volumePercent) ||
                    volumePercent <= 0.0 ||
                    volumePercent > 100.0)
                {
                    return null;
                }

                double requestedRank =
                    Math.Ceiling(volumePercent / 100.0 * totalSamples);
                if (!IsFinite(requestedRank) || requestedRank > long.MaxValue)
                    return null;

                long rank = requestedRank < 1.0
                    ? 1L
                    : (long)requestedRank;
                long cumulative = 0;

                for (int bin = histogram.Length - 1; bin >= 0; bin--)
                {
                    if (long.MaxValue - cumulative < histogram[bin])
                        return null;

                    cumulative += histogram[bin];
                    if (cumulative < rank)
                        continue;

                    if (rank == 1)
                        return maxDoseGy;

                    double lowerBinEdge = bin * binWidthGy;
                    return IsFinite(lowerBinEdge)
                        ? (double?)lowerBinEdge
                        : null;
                }

                return null;
            }

            private static bool IsFinitePositive(double value)
            {
                return IsFinite(value) && value > 0.0;
            }
        }

        private sealed class DvhHelper
        {
            private readonly PlanSetup _plan;
            private readonly double? _binWidth;
            private readonly bool _dosePresentationStateSafe;
            private readonly string _binWidthWarning;

            public DvhHelper(PlanSetup plan)
            {
                _plan = plan;
                bool restoreFailed;
                bool initializationFailed;
                _binWidth = GetBinWidthForPlanDoseUnit(
                    out restoreFailed,
                    out initializationFailed);
                _dosePresentationStateSafe = !restoreFailed;
                _binWidthWarning = restoreFailed
                    ? "DVH_DOSE_PRESENTATION_RESTORE_ERROR"
                    : initializationFailed
                        ? "DVH_BIN_WIDTH_ERROR"
                        : null;
            }

            public bool DosePresentationStateSafe
            {
                get { return _dosePresentationStateSafe; }
            }

            public DvhMetricsResult GetMetrics(Structure structure)
            {
                try
                {
                    if (structure == null || structure.IsEmpty)
                        return new DvhMetricsResult("UNAVAILABLE");

                    if (!_dosePresentationStateSafe)
                    {
                        var unsafePresentation =
                            new DvhMetricsResult("PRESENTATION_RESTORE_ERROR");
                        unsafePresentation.AddWarning(
                            "DVH_DOSE_PRESENTATION_RESTORE_ERROR");
                        return unsafePresentation;
                    }

                    if (!_binWidth.HasValue)
                    {
                        var unsupported = new DvhMetricsResult(
                            _binWidthWarning == "DVH_BIN_WIDTH_ERROR"
                                ? "ERROR"
                                : "DOSE_UNIT_UNSUPPORTED");
                        unsupported.AddWarning(
                            _binWidthWarning ??
                            "DVH_BIN_WIDTH_UNIT_UNSUPPORTED");
                        return unsupported;
                    }

                    var dvh = _plan.GetDVHCumulativeData(
                        structure,
                        DoseValuePresentation.Absolute,
                        VolumePresentation.Relative,
                        _binWidth.Value);

                    if (dvh == null || dvh.CurveData == null)
                    {
                        var unavailable = new DvhMetricsResult("UNAVAILABLE");
                        unavailable.AddWarning("DVH_CURVE_UNAVAILABLE");
                        return unavailable;
                    }

                    List<DVHPoint> curve = dvh.CurveData.ToList();
                    if (curve.Count == 0)
                    {
                        var unavailable = new DvhMetricsResult("UNAVAILABLE");
                        unavailable.AddWarning("DVH_CURVE_UNAVAILABLE");
                        return unavailable;
                    }

                    var result = new DvhMetricsResult("")
                    {
                        Coverage = IsFinite(dvh.Coverage)
                            ? (double?)dvh.Coverage
                            : null,
                        SamplingCoverage = IsFinite(dvh.SamplingCoverage)
                            ? (double?)dvh.SamplingCoverage
                            : null
                    };

                    AddCoverageWarnings(result, dvh.Coverage, dvh.SamplingCoverage);

                    if (curve.Any(point =>
                        !IsFinite(point.Volume) ||
                        point.Volume < 0.0 ||
                        point.Volume > 100.0 ||
                        !DoseToGy(point.DoseValue).HasValue))
                    {
                        result.AddWarning("DVH_CURVE_POINTS_REJECTED");
                    }

                    result.D2Gy = DoseAtVolumePercentGy(curve, 2);
                    result.D50Gy = DoseAtVolumePercentGy(curve, 50);
                    result.D60Gy = DoseAtVolumePercentGy(curve, 60);

                    int metricCount = new[]
                    {
                        result.D2Gy,
                        result.D50Gy,
                        result.D60Gy
                    }.Count(value => value.HasValue);

                    if (metricCount == 3)
                    {
                        result.Status = result.WarningCodes.Count == 0
                            ? "OK"
                            : "WARNING";
                    }
                    else if (metricCount > 0)
                    {
                        result.Status = "PARTIAL";
                        result.AddWarning("DVH_METRICS_PARTIAL");
                    }
                    else
                    {
                        result.Status = "UNAVAILABLE";
                        result.AddWarning("DVH_METRICS_UNAVAILABLE");
                    }

                    return result;
                }
                catch
                {
                    var failed = new DvhMetricsResult("ERROR");
                    failed.AddWarning("DVH_EXTRACTION_ERROR");
                    return failed;
                }
            }

            private static void AddCoverageWarnings(
                DvhMetricsResult result,
                double coverage,
                double samplingCoverage)
            {
                if (!IsFinite(coverage))
                    result.AddWarning("DVH_COVERAGE_UNAVAILABLE");
                else if (coverage < MinimumDvhCoverage)
                    result.AddWarning("DVH_COVERAGE_LOW");
                else if (coverage > 1.0)
                    result.AddWarning("DVH_COVERAGE_GT_ONE");

                if (!IsFinite(samplingCoverage))
                    result.AddWarning("DVH_SAMPLING_COVERAGE_UNAVAILABLE");
                else if (samplingCoverage < MinimumDvhCoverage)
                    result.AddWarning("DVH_SAMPLING_COVERAGE_LOW");
                else if (samplingCoverage > 1.0)
                    result.AddWarning("DVH_SAMPLING_COVERAGE_GT_ONE");
            }

            private double? GetBinWidthForPlanDoseUnit(
                out bool restoreFailed,
                out bool initializationFailed)
            {
                restoreFailed = false;
                initializationFailed = false;
                DoseValuePresentation originalPresentation;

                try
                {
                    originalPresentation = _plan.DoseValuePresentation;
                }
                catch
                {
                    initializationFailed = true;
                    return null;
                }

                bool presentationChanged =
                    originalPresentation != DoseValuePresentation.Absolute;
                double? binWidth = null;

                try
                {
                    if (presentationChanged)
                    {
                        _plan.DoseValuePresentation =
                            DoseValuePresentation.Absolute;
                    }

                    double? totalDoseBinWidth =
                        GetBinWidthForDoseValue(_plan.TotalDose);
                    if (totalDoseBinWidth.HasValue)
                    {
                        binWidth = totalDoseBinWidth;
                    }
                    else
                    {
                        double? fractionDoseBinWidth =
                            GetBinWidthForDoseValue(_plan.DosePerFraction);
                        binWidth = fractionDoseBinWidth.HasValue
                            ? fractionDoseBinWidth
                            : GetBinWidthForDoseValue(_plan.Dose.DoseMax3D);
                    }
                }
                catch
                {
                    initializationFailed = true;
                    binWidth = null;
                }
                finally
                {
                    if (presentationChanged)
                    {
                        try
                        {
                            _plan.DoseValuePresentation = originalPresentation;
                        }
                        catch
                        {
                            restoreFailed = true;
                            binWidth = null;
                        }
                    }
                }

                return binWidth;
            }

            private static double? GetBinWidthForDoseValue(DoseValue doseValue)
            {
                if (doseValue.IsUndefined())
                    return null;

                DoseValue.DoseUnit unit = doseValue.Unit;

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
                double? previousDoseGy = null;

                foreach (var point in curve)
                {
                    double? pointDoseGy = DoseToGy(point.DoseValue);
                    if (!IsFinite(point.Volume) ||
                        point.Volume < 0.0 ||
                        point.Volume > 100.0 ||
                        !pointDoseGy.HasValue)
                    {
                        previous = null;
                        previousDoseGy = null;
                        continue;
                    }

                    if (point.Volume <= volumePercent)
                    {
                        if (!previous.HasValue || !previousDoseGy.HasValue)
                        {
                            return Math.Abs(point.Volume - volumePercent) < 1e-6
                                ? pointDoseGy
                                : null;
                        }

                        double firstVolume = previous.Value.Volume;
                        double secondVolume = point.Volume;

                        if (firstVolume < volumePercent || secondVolume > volumePercent)
                            return null;

                        if (Math.Abs(secondVolume - firstVolume) < 1e-6)
                            return pointDoseGy.Value;

                        double fraction =
                            (volumePercent - firstVolume) / (secondVolume - firstVolume);
                        double interpolated =
                            previousDoseGy.Value +
                            fraction * (pointDoseGy.Value - previousDoseGy.Value);

                        return IsFinite(interpolated) ? (double?)interpolated : null;
                    }

                    previous = point;
                    previousDoseGy = pointDoseGy;
                }

                return null;
            }

            private static double? DoseToGy(DoseValue doseValue)
            {
                if (doseValue.IsUndefined() ||
                    !IsFinite(doseValue.Dose) ||
                    doseValue.Dose < 0.0)
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
