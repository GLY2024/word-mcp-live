using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Web.Script.Serialization;
using System.Xml;

[assembly: AssemblyVersion("1.0.0.0")]

namespace WordMcpLive.MathTypeBridge
{
    internal static class Program
    {
        private const uint CoinitApartmentThreaded = 0x2;
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();

        [DllImport("ole32.dll")]
        private static extern int CoInitializeEx(IntPtr reserved, uint coInit);

        [DllImport("ole32.dll")]
        private static extern void CoUninitialize();

        public static int Main()
        {
            Console.InputEncoding = new UTF8Encoding(false);
            Console.OutputEncoding = new UTF8Encoding(false);
            int hr = CoInitializeEx(IntPtr.Zero, CoinitApartmentThreaded);
            bool initialized = hr >= 0;

            try
            {
                Dictionary<string, object> request = ReadRequest();
                Dictionary<string, object> result = Execute(request, null);
                WriteResponse(true, result, null);
                return 0;
            }
            catch (BridgeException error)
            {
                WriteResponse(false, null, Error(error.Code, error.Message));
                return 1;
            }
            catch (Exception error)
            {
                WriteResponse(false, null, Error("bridge_failure", error.Message));
                return 1;
            }
            finally
            {
                if (initialized)
                {
                    CoUninitialize();
                }
            }
        }

        private static Dictionary<string, object> ReadRequest()
        {
            string input = Console.In.ReadToEnd();
            if (String.IsNullOrWhiteSpace(input))
            {
                throw new BridgeException("invalid_request", "The request body is empty.");
            }

            try
            {
                Dictionary<string, object> request = Json.Deserialize<Dictionary<string, object>>(input);
                if (request == null)
                {
                    throw new BridgeException("invalid_request", "The request must be a JSON object.");
                }
                return request;
            }
            catch (InvalidOperationException error)
            {
                throw new BridgeException("invalid_request", "Invalid request JSON: " + error.Message);
            }
        }

        internal static Dictionary<string, object> Execute(
            Dictionary<string, object> request,
            object wordApplication)
        {
            string command = RequiredString(request, "command");
            if (command == "validate_mathml")
            {
                MathMlValue value = MathMl.Parse(RequiredString(request, "mathml"));
                return Dict(
                    "mathml_sha256", value.Sha256,
                    "canonical_mathml", value.CanonicalXml
                );
            }

            if (wordApplication == null)
            {
                return AddInProxy.Execute(request);
            }

            using (WordSession session = WordSession.FromApplication(
                wordApplication,
                OptionalString(request, "filename")))
            {
                if (command == "list_equations")
                {
                    return session.ListEquations();
                }
                if (command == "get_equation")
                {
                    return session.GetEquation(RequiredString(request, "equation_id"));
                }
                if (command == "dump_equations")
                {
                    return session.DumpEquations(RequiredString(request, "output_path"));
                }
                if (command == "dump_document")
                {
                    return session.DumpDocument(RequiredString(request, "output_path"));
                }
                if (command == "delete_equation")
                {
                    return session.DeleteEquation(RequiredString(request, "equation_id"));
                }
                if (command == "probe_equation")
                {
                    return session.ProbeEquation(RequiredString(request, "equation_id"));
                }
                if (command == "replace_equation_tex")
                {
                    return session.ReplaceEquationWithTex(
                        RequiredString(request, "equation_id"),
                        RequiredString(request, "tex"),
                        RequiredString(request, "expected_mathml_sha256"));
                }
                if (command == "replace_equation")
                {
                    return session.ReplaceEquation(
                        RequiredString(request, "equation_id"),
                        RequiredString(request, "mathml"),
                        RequiredString(request, "expected_mathml_sha256")
                    );
                }
            }

            throw new BridgeException("unknown_command", "Unknown command: " + command);
        }

        private static string RequiredString(Dictionary<string, object> values, string name)
        {
            string value = OptionalString(values, name);
            if (String.IsNullOrEmpty(value))
            {
                throw new BridgeException("invalid_request", name + " is required.");
            }
            return value;
        }

        private static string OptionalString(Dictionary<string, object> values, string name)
        {
            object raw;
            if (!values.TryGetValue(name, out raw) || raw == null)
            {
                return null;
            }
            string value = raw as string;
            if (value == null)
            {
                throw new BridgeException("invalid_request", name + " must be a string or null.");
            }
            return value;
        }

        private static Dictionary<string, object> Dict(params object[] pairs)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            for (int index = 0; index < pairs.Length; index += 2)
            {
                result[(string)pairs[index]] = pairs[index + 1];
            }
            return result;
        }

        internal static Dictionary<string, object> Error(string code, string message)
        {
            return Dict("code", code, "message", message);
        }

        private static void WriteResponse(
            bool ok,
            Dictionary<string, object> result,
            Dictionary<string, object> error)
        {
            Dictionary<string, object> response = new Dictionary<string, object>();
            response["ok"] = ok;
            if (ok)
            {
                response["result"] = result;
            }
            else
            {
                response["error"] = error;
            }
            Console.Write(Json.Serialize(response));
        }

        internal static string SerializeResponse(
            bool ok,
            Dictionary<string, object> result,
            Dictionary<string, object> error)
        {
            Dictionary<string, object> response = new Dictionary<string, object>();
            response["ok"] = ok;
            response[ok ? "result" : "error"] = ok ? (object)result : error;
            return Json.Serialize(response);
        }
    }

    [ComVisible(true)]
    [Guid("C4B5A18E-3427-49D9-94D4-36C8AF8B5F61")]
    [ProgId("WordMcpLive.MathTypeAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class MathTypeAddIn : Extensibility.IDTExtensibility2
    {
        private object _application;

        public string Execute(string requestJson)
        {
            JavaScriptSerializer json = new JavaScriptSerializer();
            try
            {
                Dictionary<string, object> request =
                    json.Deserialize<Dictionary<string, object>>(requestJson);
                if (request == null)
                {
                    throw new BridgeException("invalid_request", "The request must be a JSON object.");
                }
                Dictionary<string, object> result = Program.Execute(request, _application);
                return Program.SerializeResponse(true, result, null);
            }
            catch (BridgeException error)
            {
                return Program.SerializeResponse(
                    false,
                    null,
                    Program.Error(error.Code, error.Message)
                );
            }
            catch (Exception error)
            {
                return Program.SerializeResponse(
                    false,
                    null,
                    Program.Error("bridge_failure", error.Message)
                );
            }
        }

        public void OnConnection(
            object application,
            Extensibility.ext_ConnectMode connectMode,
            object addInInstance,
            ref Array custom)
        {
            _application = application;
            dynamic addIn = addInInstance;
            addIn.Object = this;
        }

        public void OnDisconnection(
            Extensibility.ext_DisconnectMode removeMode,
            ref Array custom)
        {
            _application = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }
    }

    internal static class AddInProxy
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();

        public static Dictionary<string, object> Execute(Dictionary<string, object> request)
        {
            object wordObject = null;
            object comAddInsObject = null;
            object addInObject = null;
            object automationObject = null;
            try
            {
                wordObject = WordSession.GetActiveWordApplication();
                dynamic word = wordObject;
                comAddInsObject = word.COMAddIns;
                dynamic comAddIns = comAddInsObject;
                addInObject = comAddIns.Item("WordMcpLive.MathTypeAddIn");
                dynamic addIn = addInObject;
                if (!(bool)addIn.Connect)
                {
                    addIn.Connect = true;
                }
                automationObject = addIn.Object;
                if (automationObject == null)
                {
                    throw new BridgeException(
                        "addin_not_loaded",
                        "The MathType Word add-in is registered but did not expose its automation object."
                    );
                }

                dynamic automation = automationObject;
                string responseJson = automation.Execute(Json.Serialize(request));
                Dictionary<string, object> response =
                    Json.Deserialize<Dictionary<string, object>>(responseJson);
                if (response == null || !response.ContainsKey("ok"))
                {
                    throw new BridgeException(
                        "invalid_addin_response", "The MathType Word add-in returned invalid JSON."
                    );
                }
                if (!(bool)response["ok"])
                {
                    Dictionary<string, object> error =
                        response["error"] as Dictionary<string, object>;
                    throw new BridgeException(
                        error != null && error.ContainsKey("code")
                            ? (string)error["code"] : "addin_error",
                        error != null && error.ContainsKey("message")
                            ? (string)error["message"] : "The MathType Word add-in failed."
                    );
                }
                Dictionary<string, object> result =
                    response["result"] as Dictionary<string, object>;
                if (result == null)
                {
                    throw new BridgeException(
                        "invalid_addin_response", "The MathType Word add-in result is invalid."
                    );
                }
                return result;
            }
            catch (COMException error)
            {
                throw new BridgeException(
                    "addin_not_loaded",
                    "The MathType Word add-in is not registered or loaded. Run"
                        + " word_document_server/mathtype_bridge/install_mathtype_addin.ps1"
                        + " and restart Word. Details: " + error.Message
                );
            }
            finally
            {
                ComObjects.ReleaseAll(
                    automationObject, addInObject, comAddInsObject, wordObject);
            }
        }
    }

    internal sealed class WordSession : IDisposable
    {
        private const int RunForConversionVerb = 2;
        private readonly object _wordObject;
        private readonly object _documentObject;
        private readonly dynamic _word;
        private readonly dynamic _document;
        private readonly bool _ownsWordReference;
        private bool _disposed;

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string progId,
            out Guid clsid);

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(
            ref Guid clsid,
            IntPtr reserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object value);

        private WordSession(object word, object document, bool ownsWordReference)
        {
            _wordObject = word;
            _documentObject = document;
            _word = word;
            _document = document;
            _ownsWordReference = ownsWordReference;
        }

        public static WordSession Attach(string filename)
        {
            return Create(GetActiveWordApplication(), filename, true);
        }

        public static WordSession FromApplication(object word, string filename)
        {
            return Create(word, filename, false);
        }

        internal static object GetActiveWordApplication()
        {
            Guid wordClsid;
            int hr = CLSIDFromProgID("Word.Application", out wordClsid);
            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }

            object word;
            try
            {
                GetActiveObject(ref wordClsid, IntPtr.Zero, out word);
            }
            catch (COMException error)
            {
                throw new BridgeException(
                    "word_not_running",
                    "Microsoft Word is not running or is not accessible: " + error.Message
                );
            }

            return word;
        }

        private static WordSession Create(object word, string filename, bool ownsWordReference)
        {
            dynamic application = word;
            object document = null;
            object documentsObject = null;
            try
            {
                documentsObject = application.Documents;
                dynamic documents = documentsObject;
                if (String.IsNullOrEmpty(filename))
                {
                    if ((int)documents.Count == 0)
                    {
                        throw new BridgeException("document_not_open", "Word has no open document.");
                    }
                    document = application.ActiveDocument;
                }
                else
                {
                    string requestedFullPath = Path.IsPathRooted(filename)
                        ? Path.GetFullPath(filename)
                        : null;
                    int count = (int)documents.Count;
                    for (int index = 1; index <= count; index++)
                    {
                        object candidateObject = null;
                        try
                        {
                            candidateObject = documents[index];
                            dynamic candidate = candidateObject;
                            string candidateName = (string)candidate.Name;
                            string candidateFullName = (string)candidate.FullName;
                            if (String.Equals(
                                    candidateName,
                                    filename,
                                    StringComparison.OrdinalIgnoreCase)
                                || (requestedFullPath != null
                                    && String.Equals(
                                        Path.GetFullPath(candidateFullName),
                                        requestedFullPath,
                                        StringComparison.OrdinalIgnoreCase)))
                            {
                                document = candidateObject;
                                candidateObject = null;
                                break;
                            }
                        }
                        finally
                        {
                            ComObjects.Release(candidateObject);
                        }
                    }
                    if (document == null)
                    {
                        throw new BridgeException(
                            "document_not_open",
                            "The requested document is not open in Word: " + filename
                        );
                    }
                }

                return new WordSession(word, document, ownsWordReference);
            }
            catch
            {
                ComObjects.ReleaseAll(document, ownsWordReference ? word : null);
                throw;
            }
            finally
            {
                ComObjects.Release(documentsObject);
            }
        }

        public Dictionary<string, object> ListEquations()
        {
            using (EquationInventory inventory = EnumerateEquations())
            {
                List<Dictionary<string, object>> items =
                    new List<Dictionary<string, object>>();
                foreach (EquationReference equation in inventory.Items)
                {
                    Dictionary<string, object> metadata = equation.ToMetadata();
                    string layout;
                    string number;
                    ClassifyEquation(equation, out layout, out number);
                    metadata["layout"] = layout;
                    if (number.Length > 0)
                    {
                        metadata["number"] = number;
                    }
                    items.Add(metadata);
                }
                return ProgramDict(
                    "document", (string)_document.Name,
                    "equations", items,
                    "count", items.Count
                );
            }
        }

        public Dictionary<string, object> GetEquation(string equationId)
        {
            using (EquationInventory inventory = EnumerateEquations())
            {
                EquationReference equation = inventory.FindById(equationId);
                OleMathMl read = MathTypeOle.Read(equation.OleFormat, RunForConversionVerb);
                return ProgramDict(
                    "equation_id", equation.Id,
                    "document", (string)_document.Name,
                    "prog_id", equation.ProgId,
                    "mathml_format", read.FormatName,
                    "mathml", read.Value.OriginalXml,
                    "canonical_mathml", read.Value.CanonicalXml,
                    "mathml_sha256", read.Value.Sha256
                );
            }
        }

        public Dictionary<string, object> DumpEquations(string outputPath)
        {
            using (EquationInventory inventory = EnumerateEquations())
            {
                List<EquationReference> equations = inventory.Items;
                StringBuilder output = new StringBuilder();
                int failed = 0;
                int index = 0;
                foreach (EquationReference equation in equations)
                {
                    index++;
                    string body;
                    string sha256 = "";
                    try
                    {
                        OleMathMl read = MathTypeOle.Read(
                            equation.OleFormat, RunForConversionVerb);
                        sha256 = read.Value.Sha256;
                        body = FlattenLine(read.Value.CanonicalXml);
                    }
                    catch (Exception error)
                    {
                        failed++;
                        body = "error: " + FlattenLine(error.Message);
                    }
                    string layout;
                    string number;
                    ClassifyEquation(equation, out layout, out number);
                    output.AppendLine(String.Format(
                        "[eq {0}] layout={1}{2} id={3} sha256={4}",
                        index,
                        layout,
                        number.Length > 0 ? " number=" + number : "",
                        equation.Id,
                        sha256));
                    output.AppendLine("ctx: " + ContextText(equation));
                    output.AppendLine(body);
                    output.AppendLine();
                }

                string fullPath = Path.GetFullPath(outputPath);
                string header = String.Format(
                    "# {0} | {1} equations | {2} failed{3}",
                    (string)_document.Name,
                    equations.Count,
                    failed,
                    Environment.NewLine);
                File.WriteAllText(fullPath, header + output, new UTF8Encoding(false));
                return ProgramDict(
                    "document", (string)_document.Name,
                    "count", equations.Count,
                    "failed", failed,
                    "output_path", fullPath
                );
            }
        }

        public Dictionary<string, object> DumpDocument(string outputPath)
        {
            using (EquationInventory inventory = EnumerateEquations())
            {
                List<EquationReference> equations = inventory.Items.FindAll(
                    delegate(EquationReference equation) { return equation.StoryType == 1; });
                equations.Sort(delegate(EquationReference left, EquationReference right)
                {
                    int byStart = left.RangeStart.CompareTo(right.RangeStart);
                    return byStart != 0 ? byStart : left.RangeEnd.CompareTo(right.RangeEnd);
                });

                int storyStart;
                int storyEnd;
                GetMainStoryBounds(out storyStart, out storyEnd);
                int cursor = storyStart;
                StringBuilder output = new StringBuilder();
                int index = 0;
                foreach (EquationReference equation in equations)
                {
                    if (equation.RangeStart < cursor || equation.RangeEnd > storyEnd)
                    {
                        throw new BridgeException(
                            "invalid_equation_range",
                            "A MathType equation range overlaps another equation or lies outside the main story."
                        );
                    }
                    output.Append(ReadDocumentRangeText(cursor, equation.RangeStart));
                    index++;
                    string markerLayout;
                    string markerNumber;
                    ClassifyEquation(equation, out markerLayout, out markerNumber);
                    output.Append(markerNumber.Length > 0
                        ? String.Format("[eq {0} {1}]", index, markerNumber)
                        : String.Format("[eq {0}]", index));
                    cursor = equation.RangeEnd;
                }
                output.Append(ReadDocumentRangeText(cursor, storyEnd));
                output.AppendLine();
                output.AppendLine();
                output.AppendLine("=== MathType equations ===");

                int failed = 0;
                index = 0;
                foreach (EquationReference equation in equations)
                {
                    index++;
                    string layout;
                    string number;
                    ClassifyEquation(equation, out layout, out number);
                    string body;
                    string sha256 = "";
                    try
                    {
                        OleMathMl read = MathTypeOle.Read(
                            equation.OleFormat, RunForConversionVerb);
                        sha256 = read.Value.Sha256;
                        body = FlattenLine(read.Value.CanonicalXml);
                    }
                    catch (Exception error)
                    {
                        failed++;
                        body = "error: " + FlattenLine(error.Message);
                    }
                    output.AppendLine(String.Format(
                        "[eq {0}] layout={1}{2} id={3} sha256={4}",
                        index,
                        layout,
                        number.Length > 0 ? " number=" + number : "",
                        equation.Id,
                        sha256));
                    output.AppendLine(body);
                    output.AppendLine();
                }

                string fullPath = Path.GetFullPath(outputPath);
                File.WriteAllText(fullPath, output.ToString(), new UTF8Encoding(false));
                return ProgramDict(
                    "document", (string)_document.Name,
                    "count", equations.Count,
                    "failed", failed,
                    "output_path", fullPath
                );
            }
        }

        private void GetMainStoryBounds(out int start, out int end)
        {
            object storyRangesObject = null;
            object mainStoryObject = null;
            try
            {
                storyRangesObject = _document.StoryRanges;
                dynamic storyRanges = storyRangesObject;
                mainStoryObject = storyRanges[1]; // wdMainTextStory
                dynamic mainStory = mainStoryObject;
                start = (int)mainStory.Start;
                end = (int)mainStory.End;
            }
            finally
            {
                ComObjects.ReleaseAll(mainStoryObject, storyRangesObject);
            }
        }

        private string ReadDocumentRangeText(int start, int end)
        {
            object rangeObject = null;
            try
            {
                rangeObject = _document.Range(start, end);
                dynamic range = rangeObject;
                return (string)range.Text;
            }
            finally
            {
                ComObjects.Release(rangeObject);
            }
        }

        private int DocumentContentEnd()
        {
            object contentObject = null;
            try
            {
                contentObject = _document.Content;
                dynamic content = contentObject;
                return (int)content.End;
            }
            finally
            {
                ComObjects.Release(contentObject);
            }
        }

        public Dictionary<string, object> DeleteEquation(string equationId)
        {
            using (EquationInventory inventory = EnumerateEquations())
            {
                EquationReference equation = inventory.FindById(equationId);
                object undoRecordObject = _word.UndoRecord;
                dynamic undoRecord = undoRecordObject;
                bool undoStarted = false;
                try
                {
                    undoRecord.StartCustomRecord("Delete MathType equation");
                    undoStarted = true;
                    equation.Shape.Delete();
                    return ProgramDict(
                        "equation_id", equation.Id,
                        "document", (string)_document.Name,
                        "deleted", true
                    );
                }
                finally
                {
                    try
                    {
                        if (undoStarted)
                        {
                            undoRecord.EndCustomRecord();
                        }
                    }
                    finally
                    {
                        ComObjects.Release(undoRecordObject);
                    }
                }
            }
        }

        public Dictionary<string, object> ProbeEquation(string equationId)
        {
            using (EquationInventory inventory = EnumerateEquations())
            {
                EquationReference equation = inventory.FindById(equationId);
                return MathTypeOle.Probe(equation.OleFormat, RunForConversionVerb);
            }
        }

        public Dictionary<string, object> ReplaceEquationWithTex(
            string equationId, string tex, string expectedSha256)
        {
            ValidateSha256(expectedSha256);
            using (EquationInventory originalInventory = EnumerateEquations())
            {
                EquationReference equation = originalInventory.FindById(equationId);
                if (equation.Kind != "inline")
                {
                    throw new BridgeException(
                        "inline_only",
                        "TeX replacement currently supports inline equations only."
                    );
                }

                OleMathMl current = MathTypeOle.Read(
                    equation.OleFormat, RunForConversionVerb);
                if (!String.Equals(
                    current.Value.Sha256,
                    expectedSha256,
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new BridgeException(
                        "equation_changed",
                        "The equation changed after it was read; read it again before replacing it."
                    );
                }

                object selectionObject = null;
                object selectionRangeObject = null;
                object originalSelectionObject = null;
                object originalDocumentObject = null;
                object undoRecordObject = null;
                object shapeRangeObject = null;
                object insertRangeObject = null;
                bool undoStarted = false;
                bool mutated = false;
                int oldAlerts = (int)_word.DisplayAlerts;
                bool oldScreenUpdating = (bool)_word.ScreenUpdating;
                try
                {
                    selectionObject = _word.Selection;
                    dynamic selection = selectionObject;
                    selectionRangeObject = selection.Range;
                    dynamic selectionRange = selectionRangeObject;
                    originalSelectionObject = selectionRange.Duplicate;
                    originalDocumentObject = _word.ActiveDocument;

                    undoRecordObject = _word.UndoRecord;
                    dynamic undoRecord = undoRecordObject;

                    _word.DisplayAlerts = 0;
                    _word.ScreenUpdating = false;
                    _document.Activate();
                    undoRecord.StartCustomRecord("Replace MathType equation (TeX)");
                    undoStarted = true;

                    dynamic shape = equation.Shape;
                    shapeRangeObject = shape.Range;
                    dynamic shapeRange = shapeRangeObject;
                    insertRangeObject = shapeRange.Duplicate;
                    dynamic insertRange = insertRangeObject;
                    shape.Delete();
                    mutated = true;

                    string marked = "$" + tex + "$";
                    insertRange.Text = marked;
                    insertRange.Select();
                    _word.Run("MTCommand_TexToggle");

                    using (EquationInventory replacementInventory = EnumerateEquations())
                    {
                        EquationReference replacement = replacementInventory.FindAt(equation);
                        if (replacement == null)
                        {
                            throw new BridgeException(
                                "tex_conversion_failed",
                                "MathType Toggle TeX did not produce an equation; the original"
                                    + " equation was restored. Check the TeX syntax."
                            );
                        }

                        OleMathMl read = MathTypeOle.Read(
                            replacement.OleFormat, RunForConversionVerb);
                        Dictionary<string, object> result = ProgramDict(
                            "equation_id", replacement.Id,
                            "replaced_equation_id", equationId,
                            "document", (string)_document.Name,
                            "mathml", read.Value.CanonicalXml,
                            "mathml_sha256", read.Value.Sha256
                        );
                        undoRecord.EndCustomRecord();
                        undoStarted = false;
                        return result;
                    }
                }
                catch (Exception original)
                {
                    if (!mutated)
                    {
                        throw;
                    }
                    if (undoStarted)
                    {
                        try
                        {
                            dynamic undoRecord = undoRecordObject;
                            undoRecord.EndCustomRecord();
                        }
                        catch (COMException)
                        {
                            // Undo is still attempted and verified below.
                        }
                        undoStarted = false;
                    }
                    bool restored = false;
                    try
                    {
                        _document.Undo(1);
                        using (EquationInventory restoredInventory = EnumerateEquations())
                        {
                            EquationReference candidate = restoredInventory.FindAt(equation);
                            if (candidate != null)
                            {
                                OleMathMl after = MathTypeOle.Read(
                                    candidate.OleFormat, RunForConversionVerb);
                                restored = String.Equals(
                                    after.Value.Sha256,
                                    expectedSha256,
                                    StringComparison.OrdinalIgnoreCase);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        restored = false;
                    }
                    if (!restored)
                    {
                        throw new BridgeException(
                            "rollback_failed",
                            "TeX replacement failed and the automatic rollback could not be"
                                + " verified; inspect the document and restore it manually."
                                + " Original error: " + original.Message
                        );
                    }
                    throw;
                }
                finally
                {
                    try
                    {
                        if (undoStarted)
                        {
                            dynamic undoRecord = undoRecordObject;
                            undoRecord.EndCustomRecord();
                        }
                    }
                    finally
                    {
                        try
                        {
                            try
                            {
                                _word.ScreenUpdating = oldScreenUpdating;
                            }
                            finally
                            {
                                _word.DisplayAlerts = oldAlerts;
                            }
                            try
                            {
                                if (originalDocumentObject != null)
                                {
                                    dynamic originalDocument = originalDocumentObject;
                                    originalDocument.Activate();
                                }
                                if (originalSelectionObject != null)
                                {
                                    dynamic originalSelection = originalSelectionObject;
                                    originalSelection.Select();
                                }
                            }
                            catch (COMException)
                            {
                                // Selection restoration does not affect document contents.
                            }
                        }
                        finally
                        {
                            ComObjects.ReleaseAll(
                                insertRangeObject,
                                shapeRangeObject,
                                originalSelectionObject,
                                selectionRangeObject,
                                selectionObject,
                                originalDocumentObject,
                                undoRecordObject);
                        }
                    }
                }
            }
        }

        // Layout classification: inline text, own display paragraph (optionally with a
        // MathType right-side number via MACROBUTTON MTPlaceRef + SEQ MTEqn), or table cell.
        private void ClassifyEquation(
            EquationReference equation, out string layout, out string number)
        {
            layout = equation.Kind == "inline" ? "inline" : "floating";
            number = "";
            if (equation.Kind != "inline")
            {
                return;
            }
            object anchorObject = null;
            object paragraphsObject = null;
            object paragraphObject = null;
            object paragraphRangeObject = null;
            object fieldsObject = null;
            try
            {
                dynamic shape = equation.Shape;
                anchorObject = shape.Range;
                dynamic anchor = anchorObject;
                if ((bool)anchor.Information[12]) // wdWithInTable
                {
                    layout = "table";
                    return;
                }
                paragraphsObject = anchor.Paragraphs;
                dynamic paragraphs = paragraphsObject;
                paragraphObject = paragraphs[1];
                dynamic paragraph = paragraphObject;
                paragraphRangeObject = paragraph.Range;
                dynamic para = paragraphRangeObject;
                string stripped = FlattenLine((string)para.Text).Replace(" ", "");
                if (stripped.Length != 0)
                {
                    return;
                }
                layout = "display";
                fieldsObject = para.Fields;
                dynamic fields = fieldsObject;
                int fieldCount = (int)fields.Count;
                for (int index = 1; index <= fieldCount; index++)
                {
                    object fieldObject = null;
                    object codeObject = null;
                    object resultObject = null;
                    try
                    {
                        fieldObject = fields[index];
                        dynamic field = fieldObject;
                        int fieldType = (int)field.Type;
                        if (fieldType == 51) // wdFieldMacroButton: MTPlaceRef equation number
                        {
                            layout = "display_numbered";
                        }
                        if (fieldType == 12) // wdFieldSequence
                        {
                            codeObject = field.Code;
                            dynamic codeRange = codeObject;
                            string code = (string)codeRange.Text;
                            if (code.Contains("SEQ MTEqn") && code.Contains("\\c"))
                            {
                                resultObject = field.Result;
                                dynamic resultRange = resultObject;
                                string value = ((string)resultRange.Text).Trim();
                                if (value.Length > 0)
                                {
                                    number = "(" + value + ")";
                                }
                            }
                        }
                    }
                    finally
                    {
                        ComObjects.ReleaseAll(
                            resultObject, codeObject, fieldObject);
                    }
                }
            }
            catch (Exception)
            {
                // Classification is best effort; fall back to what we already have.
            }
            finally
            {
                ComObjects.ReleaseAll(
                    fieldsObject,
                    paragraphRangeObject,
                    paragraphObject,
                    paragraphsObject,
                    anchorObject);
            }
        }

        private string ContextText(EquationReference equation)
        {
            // ponytail: main story only; footnote/textbox context left empty until needed
            if (equation.StoryType != 1)
            {
                return "";
            }
            try
            {
                int start = equation.RangeStart;
                int documentEnd = DocumentContentEnd();
                string left = ReadDocumentRangeText(Math.Max(0, start - 60), start);
                string right = ReadDocumentRangeText(
                    Math.Min(documentEnd, start + 1),
                    Math.Min(documentEnd, start + 61));
                // Field-code text of neighbouring equations reads as "EMBED Equation.DSMT4".
                return (FlattenLine(left) + " <eq> " + FlattenLine(right))
                    .Replace("EMBED Equation.DSMT4", "[eq]");
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static string FlattenLine(string text)
        {
            if (String.IsNullOrEmpty(text))
            {
                return "";
            }
            StringBuilder result = new StringBuilder(text.Length);
            foreach (char character in text)
            {
                result.Append(character < ' ' ? ' ' : character);
            }
            return result.ToString().Trim();
        }

        public Dictionary<string, object> ReplaceEquation(
            string equationId,
            string mathml,
            string expectedSha256)
        {
            ValidateSha256(expectedSha256);
            MathMlValue replacement = MathMl.Parse(mathml);
            string resultId;
            using (EquationInventory initialInventory = EnumerateEquations())
            {
                resultId = initialInventory.FindById(equationId).Id;
            }

            object undoRecordObject = _word.UndoRecord;
            dynamic undoRecord = undoRecordObject;
            bool undoStarted = false;
            bool writeStarted = false;
            bool undoRequired = false;
            try
            {
                undoRecord.StartCustomRecord("Replace MathType equation");
                undoStarted = true;
                using (EquationInventory writableInventory = EnumerateEquations())
                {
                    EquationReference writableEquation =
                        writableInventory.FindById(equationId);
                    MathTypeOle.ReplaceIfUnchanged(
                        writableEquation.OleFormat,
                        RunForConversionVerb,
                        expectedSha256,
                        replacement.OriginalXml,
                        ref writeStarted
                    );
                }

                using (EquationInventory writtenInventory = EnumerateEquations())
                {
                    EquationReference writtenEquation = writtenInventory.FindById(equationId);
                    OleMathMl verified = MathTypeOle.Read(
                        writtenEquation.OleFormat,
                        RunForConversionVerb
                    );
                    if (!String.Equals(
                        replacement.Sha256,
                        verified.Value.Sha256,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        undoRequired = true;
                        throw new BridgeException(
                            "verification_failed",
                            "MathType did not save the requested MathML exactly; the Word change was undone."
                        );
                    }

                    Dictionary<string, object> result = ProgramDict(
                        "equation_id", resultId,
                        "document", (string)_document.Name,
                        "mathml_format", verified.FormatName,
                        "mathml", verified.Value.OriginalXml,
                        "canonical_mathml", verified.Value.CanonicalXml,
                        "mathml_sha256", verified.Value.Sha256,
                        "verified", true
                    );
                    undoRecord.EndCustomRecord();
                    undoStarted = false;
                    return result;
                }
            }
            catch
            {
                if (writeStarted)
                {
                    undoRequired = true;
                }
                throw;
            }
            finally
            {
                try
                {
                    try
                    {
                        if (undoStarted)
                        {
                            undoRecord.EndCustomRecord();
                        }
                    }
                    finally
                    {
                        if (undoRequired)
                        {
                            _document.Undo(1);
                        }
                    }
                }
                finally
                {
                    ComObjects.Release(undoRecordObject);
                }
            }
        }

        private EquationInventory EnumerateEquations()
        {
            List<EquationReference> equations = new List<EquationReference>();
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            object storyRangesObject = null;
            try
            {
                storyRangesObject = _document.StoryRanges;
                dynamic storyRanges = storyRangesObject;
                for (int storyType = 1; storyType <= 17; storyType++)
                {
                    object storyRangeObject = null;
                    try
                    {
                        storyRangeObject = storyRanges[storyType];
                    }
                    catch (COMException)
                    {
                        continue;
                    }

                    int chain = 0;
                    while (storyRangeObject != null)
                    {
                        object currentStoryObject = storyRangeObject;
                        object nextStoryObject = null;
                        storyRangeObject = null;
                        try
                        {
                            AddInlineEquations(
                                currentStoryObject, chain, equations, seen);
                            AddFloatingEquations(
                                currentStoryObject, chain, equations, seen);
                            dynamic currentStory = currentStoryObject;
                            try
                            {
                                nextStoryObject = currentStory.NextStoryRange;
                            }
                            catch (COMException)
                            {
                                nextStoryObject = null;
                            }
                        }
                        finally
                        {
                            ComObjects.Release(currentStoryObject);
                        }
                        storyRangeObject = nextStoryObject;
                        chain++;
                    }
                }
                return new EquationInventory(equations);
            }
            catch (Exception original)
            {
                try
                {
                    new EquationInventory(equations).Dispose();
                }
                catch (Exception cleanup)
                {
                    throw new AggregateException(
                        "Equation enumeration and COM cleanup both failed.",
                        original,
                        cleanup);
                }
                throw;
            }
            finally
            {
                ComObjects.Release(storyRangesObject);
            }
        }

        private static void AddInlineEquations(
            object rangeObject,
            int chain,
            List<EquationReference> equations,
            HashSet<string> seen)
        {
            object inlineShapesObject = null;
            try
            {
                dynamic range = rangeObject;
                inlineShapesObject = range.InlineShapes;
                dynamic inlineShapes = inlineShapesObject;
                int count = (int)inlineShapes.Count;
                for (int index = 1; index <= count; index++)
                {
                    object shapeObject = null;
                    object anchorObject = null;
                    EquationReference equation = null;
                    try
                    {
                        shapeObject = inlineShapes[index];
                        dynamic shape = shapeObject;
                        anchorObject = shape.Range;
                        equation = CreateEquation(
                            "inline", shapeObject, anchorObject, chain, index);
                        if (equation != null)
                        {
                            shapeObject = null;
                            if (seen.Add(equation.Id))
                            {
                                equations.Add(equation);
                                equation = null;
                            }
                        }
                    }
                    catch (COMException)
                    {
                        // A non-OLE inline shape has no usable OLEFormat.
                    }
                    finally
                    {
                        try
                        {
                            if (equation != null)
                            {
                                equation.Dispose();
                            }
                        }
                        finally
                        {
                            ComObjects.ReleaseAll(anchorObject, shapeObject);
                        }
                    }
                }
            }
            finally
            {
                ComObjects.Release(inlineShapesObject);
            }
        }

        private static void AddFloatingEquations(
            object rangeObject,
            int chain,
            List<EquationReference> equations,
            HashSet<string> seen)
        {
            object shapesObject = null;
            try
            {
                dynamic range = rangeObject;
                shapesObject = range.ShapeRange;
                if (shapesObject == null)
                {
                    return;
                }
                dynamic shapes = shapesObject;
                int count = (int)shapes.Count;
                for (int index = 1; index <= count; index++)
                {
                    object shapeObject = null;
                    object anchorObject = null;
                    EquationReference equation = null;
                    try
                    {
                        shapeObject = shapes[index];
                        dynamic shape = shapeObject;
                        anchorObject = shape.Anchor;
                        equation = CreateEquation(
                            "floating", shapeObject, anchorObject, chain, index);
                        if (equation != null)
                        {
                            shapeObject = null;
                            if (seen.Add(equation.Id))
                            {
                                equations.Add(equation);
                                equation = null;
                            }
                        }
                    }
                    catch (COMException)
                    {
                        // A non-OLE floating shape has no usable OLEFormat.
                    }
                    finally
                    {
                        try
                        {
                            if (equation != null)
                            {
                                equation.Dispose();
                            }
                        }
                        finally
                        {
                            ComObjects.ReleaseAll(anchorObject, shapeObject);
                        }
                    }
                }
            }
            catch (COMException)
            {
                // Word throws when the range has no floating shapes.
            }
            finally
            {
                ComObjects.Release(shapesObject);
            }
        }

        // A non-null result owns shapeObject and its OLEFormat; the caller owns anchorObject.
        private static EquationReference CreateEquation(
            string kind,
            object shapeObject,
            object anchorObject,
            int chain,
            int ordinal)
        {
            object oleFormatObject = null;
            try
            {
                dynamic shape = shapeObject;
                oleFormatObject = shape.OLEFormat;
                if (oleFormatObject == null)
                {
                    // Pictures and OMML shapes expose a null OLEFormat.
                    return null;
                }
                dynamic oleFormat = oleFormatObject;
                string progId = (string)oleFormat.ProgID;
                if (!IsMathTypeProgId(progId))
                {
                    return null;
                }

                dynamic anchor = anchorObject;
                int storyType = (int)anchor.StoryType;
                int rangeStart = (int)anchor.Start;
                int rangeEnd = (int)anchor.End;
                string id = String.Format(
                    "ole:{0}:{1}.{2}:{3}:{4}:{5}",
                    kind,
                    storyType,
                    chain,
                    rangeStart,
                    ordinal,
                    progId
                );
                EquationReference equation = new EquationReference(
                    id,
                    kind,
                    storyType,
                    chain,
                    rangeStart,
                    rangeEnd,
                    ordinal,
                    progId,
                    oleFormatObject,
                    shapeObject
                );
                oleFormatObject = null;
                return equation;
            }
            finally
            {
                ComObjects.Release(oleFormatObject);
            }
        }

        private static bool IsMathTypeProgId(string progId)
        {
            return String.Equals(progId, "Equation.DSMT4", StringComparison.OrdinalIgnoreCase)
                || String.Equals(progId, "Equation.DSMT36", StringComparison.OrdinalIgnoreCase);
        }

        private static void ValidateSha256(string value)
        {
            if (value.Length != 64)
            {
                throw new BridgeException(
                    "invalid_request", "expected_mathml_sha256 must contain 64 hexadecimal characters."
                );
            }
            foreach (char character in value)
            {
                if (!Uri.IsHexDigit(character))
                {
                    throw new BridgeException(
                        "invalid_request", "expected_mathml_sha256 must contain 64 hexadecimal characters."
                    );
                }
            }
        }

        private static Dictionary<string, object> ProgramDict(params object[] pairs)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            for (int index = 0; index < pairs.Length; index += 2)
            {
                result[(string)pairs[index]] = pairs[index + 1];
            }
            return result;
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            ComObjects.ReleaseAll(
                _documentObject, _ownsWordReference ? _wordObject : null);
        }
    }

    internal sealed class EquationInventory : IDisposable
    {
        private List<EquationReference> _items;

        public EquationInventory(List<EquationReference> items)
        {
            _items = items;
        }

        public List<EquationReference> Items
        {
            get { return _items; }
        }

        public EquationReference FindById(string equationId)
        {
            EquationReference equation = _items.Find(
                delegate(EquationReference candidate)
                {
                    return candidate.Id == equationId;
                });
            if (equation == null)
            {
                throw new BridgeException(
                    "equation_not_found", "MathType equation not found: " + equationId);
            }
            return equation;
        }

        public EquationReference FindAt(EquationReference original)
        {
            return _items.Find(delegate(EquationReference candidate)
            {
                return candidate.Kind == original.Kind
                    && candidate.StoryType == original.StoryType
                    && candidate.Chain == original.Chain
                    && candidate.RangeStart == original.RangeStart;
            });
        }

        public void Dispose()
        {
            List<EquationReference> items = _items;
            _items = null;
            if (items == null)
            {
                return;
            }
            List<Exception> errors = new List<Exception>();
            foreach (EquationReference equation in items)
            {
                try
                {
                    equation.Dispose();
                }
                catch (Exception error)
                {
                    errors.Add(error);
                }
            }
            if (errors.Count > 0)
            {
                throw new AggregateException(
                    "One or more MathType COM references could not be released.",
                    errors);
            }
        }
    }

    internal sealed class EquationReference : IDisposable
    {
        public readonly string Id;
        public readonly string Kind;
        public readonly int StoryType;
        public readonly int Chain;
        public readonly int RangeStart;
        public readonly int RangeEnd;
        public readonly int Ordinal;
        public readonly string ProgId;
        private object _oleFormatObject;
        private object _shapeObject;

        public dynamic OleFormat
        {
            get { return _oleFormatObject; }
        }

        public dynamic Shape
        {
            get { return _shapeObject; }
        }

        public EquationReference(
            string id,
            string kind,
            int storyType,
            int chain,
            int rangeStart,
            int rangeEnd,
            int ordinal,
            string progId,
            object oleFormat,
            object shape)
        {
            Id = id;
            Kind = kind;
            StoryType = storyType;
            Chain = chain;
            RangeStart = rangeStart;
            RangeEnd = rangeEnd;
            Ordinal = ordinal;
            ProgId = progId;
            _oleFormatObject = oleFormat;
            _shapeObject = shape;
        }

        public Dictionary<string, object> ToMetadata()
        {
            return new Dictionary<string, object>
            {
                { "equation_id", Id },
                { "kind", Kind },
                { "story_type", StoryType },
                { "range_start", RangeStart },
                { "ordinal", Ordinal },
                { "prog_id", ProgId }
            };
        }

        public void Dispose()
        {
            object oleFormatObject = _oleFormatObject;
            object shapeObject = _shapeObject;
            _oleFormatObject = null;
            _shapeObject = null;
            ComObjects.ReleaseAll(oleFormatObject, shapeObject);
        }
    }

    internal static class ComObjects
    {
        public static void Release(object value)
        {
            // Release one acquired reference without invalidating other users of a shared RCW.
            try
            {
                if (value != null && Marshal.IsComObject(value))
                {
                    Marshal.ReleaseComObject(value);
                }
            }
            catch (InvalidComObjectException)
            {
                // A disconnected RCW has already released its COM reference.
            }
        }

        public static void ReleaseAll(params object[] values)
        {
            List<Exception> errors = new List<Exception>();
            foreach (object value in values)
            {
                try
                {
                    Release(value);
                }
                catch (Exception error)
                {
                    errors.Add(error);
                }
            }
            if (errors.Count > 0)
            {
                throw new AggregateException(
                    "One or more COM references could not be released.", errors);
            }
        }
    }

    internal static class MathTypeOle
    {
        private const int DvAspectContent = 1;
        private const uint OleCloseNoSave = 1;
        private static readonly string[] MathMlFormats =
        {
            "MathML",
            "MathML Presentation",
            "application/mathml+xml"
        };

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint RegisterClipboardFormat(string format);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClipboardFormatName(
            uint format, StringBuilder name, int maxCount);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalLock(IntPtr memory);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GlobalUnlock(IntPtr memory);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern UIntPtr GlobalSize(IntPtr memory);

        [DllImport("ole32.dll")]
        private static extern void ReleaseStgMedium(ref STGMEDIUM medium);

        public static OleMathMl Read(dynamic oleFormat, int verb)
        {
            object oleObject = null;
            try
            {
                oleFormat.DoVerb(verb);
                oleObject = oleFormat.Object;
                IDataObject dataObject = oleObject as IDataObject;
                if (dataObject == null)
                {
                    throw new BridgeException(
                        "data_object_unavailable",
                        "The activated MathType equation does not expose OLE IDataObject."
                    );
                }

                FormatSelection selection = SelectMathMlFormat(dataObject);
                string mathml = ReadHGlobal(dataObject, selection.Format);
                return new OleMathMl(selection.Name, MathMl.Parse(mathml));
            }
            finally
            {
                CloseOleObject(oleObject, OleCloseNoSave);
            }
        }

        public static void ReplaceIfUnchanged(
            dynamic oleFormat,
            int verb,
            string expectedSha256,
            string mathml,
            ref bool writeStarted)
        {
            object oleObject = null;
            try
            {
                oleFormat.DoVerb(verb);
                oleObject = oleFormat.Object;
                IDataObject dataObject = oleObject as IDataObject;
                if (dataObject == null)
                {
                    throw new BridgeException(
                        "data_object_unavailable",
                        "The activated MathType equation does not expose OLE IDataObject."
                    );
                }

                FormatSelection readable = SelectMathMlFormat(dataObject);
                MathMlValue current = MathMl.Parse(ReadHGlobal(dataObject, readable.Format));
                if (!String.Equals(
                    current.Sha256,
                    expectedSha256,
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new BridgeException(
                        "equation_changed",
                        "The equation changed after it was read; read it again before replacing it."
                    );
                }

                List<string> candidates = new List<string>();
                candidates.Add(readable.Name);
                foreach (string candidate in MathMlFormats)
                {
                    if (!candidates.Contains(candidate))
                    {
                        candidates.Add(candidate);
                    }
                }

                COMException lastFormatError = null;
                foreach (string candidate in candidates)
                {
                    FORMATETC format = CreateFormat(candidate);
                    try
                    {
                        WriteHGlobal(dataObject, format, mathml);
                        writeStarted = true;
                        return;
                    }
                    catch (COMException error)
                    {
                        if (unchecked((uint)error.ErrorCode) != 0x80040064)
                        {
                            throw;
                        }
                        lastFormatError = error;
                    }
                }
                FORMATETC textFormat = CreateFormat("MathML");
                textFormat.cfFormat = 1;
                try
                {
                    WriteHGlobal(dataObject, textFormat, mathml);
                    writeStarted = true;
                    return;
                }
                catch (COMException error)
                {
                    lastFormatError = error;
                }
                throw new BridgeException(
                    "mathml_set_format_unavailable",
                    "MathType rejected all registered MathML OLE formats: "
                        + lastFormatError.Message
                );
            }
            finally
            {
                CloseOleObject(oleObject, OleCloseNoSave);
            }
        }

        public static Dictionary<string, object> Probe(dynamic oleFormat, int verb)
        {
            object oleObject = null;
            try
            {
                oleFormat.DoVerb(verb);
                oleObject = oleFormat.Object;
                IDataObject dataObject = oleObject as IDataObject;
                if (dataObject == null)
                {
                    throw new BridgeException(
                        "data_object_unavailable",
                        "The activated MathType equation does not expose OLE IDataObject."
                    );
                }
                return new Dictionary<string, object>
                {
                    { "get_formats", EnumerateFormats(dataObject, DATADIR.DATADIR_GET) },
                    { "set_formats", EnumerateFormats(dataObject, DATADIR.DATADIR_SET) }
                };
            }
            finally
            {
                CloseOleObject(oleObject, OleCloseNoSave);
            }
        }

        private static List<Dictionary<string, object>> EnumerateFormats(
            IDataObject dataObject, DATADIR direction)
        {
            List<Dictionary<string, object>> results = new List<Dictionary<string, object>>();
            IEnumFORMATETC formats = null;
            try
            {
                formats = dataObject.EnumFormatEtc(direction);
                if (formats == null)
                {
                    return results;
                }

                FORMATETC[] buffer = new FORMATETC[1];
                int[] fetched = new int[1];
                while (formats.Next(1, buffer, fetched) == 0 && fetched[0] == 1)
                {
                    results.Add(new Dictionary<string, object>
                    {
                        { "format", ClipboardFormatName(buffer[0].cfFormat) },
                        { "cf", (int)unchecked((ushort)buffer[0].cfFormat) },
                        { "tymed", buffer[0].tymed.ToString() },
                        { "aspect", (int)buffer[0].dwAspect },
                        { "lindex", buffer[0].lindex }
                    });
                }
                return results;
            }
            catch (COMException error)
            {
                results.Add(new Dictionary<string, object>
                {
                    { "error", "EnumFormatEtc failed: 0x" + error.ErrorCode.ToString("X8") }
                });
                return results;
            }
            finally
            {
                ComObjects.Release(formats);
            }
        }

        private static string ClipboardFormatName(short cfFormat)
        {
            uint cf = unchecked((ushort)cfFormat);
            if (cf < 0xC000)
            {
                return "standard:" + cf;
            }
            StringBuilder name = new StringBuilder(256);
            int length = GetClipboardFormatName(cf, name, name.Capacity);
            return length > 0 ? name.ToString() : "unknown:" + cf;
        }

        private static FormatSelection SelectMathMlFormat(IDataObject dataObject)
        {
            foreach (string name in MathMlFormats)
            {
                FORMATETC format = CreateFormat(name);
                if (dataObject.QueryGetData(ref format) == 0)
                {
                    return new FormatSelection(name, format);
                }
            }
            throw new BridgeException(
                "mathml_format_unavailable",
                "The MathType equation does not advertise a supported MathML data format."
            );
        }

        private static FORMATETC CreateFormat(string name)
        {
            uint clipboardFormat = RegisterClipboardFormat(name);
            if (clipboardFormat == 0)
            {
                throw new BridgeException(
                    "clipboard_format_registration_failed",
                    "Windows could not register the MathML data format: " + name
                );
            }
            FORMATETC format = new FORMATETC();
            format.cfFormat = unchecked((short)clipboardFormat);
            format.dwAspect = (DVASPECT)DvAspectContent;
            format.lindex = -1;
            format.ptd = IntPtr.Zero;
            format.tymed = TYMED.TYMED_HGLOBAL;
            return format;
        }

        private static string ReadHGlobal(IDataObject dataObject, FORMATETC format)
        {
            STGMEDIUM medium;
            dataObject.GetData(ref format, out medium);
            try
            {
                if (medium.tymed != TYMED.TYMED_HGLOBAL || medium.unionmember == IntPtr.Zero)
                {
                    throw new BridgeException(
                        "invalid_mathml_medium", "MathType returned MathML in an unsupported OLE medium."
                    );
                }
                ulong size = GlobalSize(medium.unionmember).ToUInt64();
                if (size == 0 || size > Int32.MaxValue)
                {
                    throw new BridgeException(
                        "invalid_mathml_medium", "MathType returned an invalid MathML buffer size."
                    );
                }
                IntPtr pointer = GlobalLock(medium.unionmember);
                if (pointer == IntPtr.Zero)
                {
                    throw new BridgeException(
                        "invalid_mathml_medium", "Windows could not lock the MathML OLE buffer."
                    );
                }
                try
                {
                    byte[] bytes = new byte[(int)size];
                    Marshal.Copy(pointer, bytes, 0, bytes.Length);
                    int length = Array.IndexOf(bytes, (byte)0);
                    if (length < 0)
                    {
                        length = bytes.Length;
                    }
                    return new UTF8Encoding(false, true).GetString(bytes, 0, length);
                }
                finally
                {
                    GlobalUnlock(medium.unionmember);
                }
            }
            finally
            {
                ReleaseStgMedium(ref medium);
            }
        }

        private static void WriteHGlobal(IDataObject dataObject, FORMATETC format, string mathml)
        {
            byte[] content = new UTF8Encoding(false, true).GetBytes(mathml);
            IntPtr memory = Marshal.AllocHGlobal(content.Length + 1);
            if (memory == IntPtr.Zero)
            {
                throw new OutOfMemoryException("Windows could not allocate the MathML OLE buffer.");
            }

            try
            {
                Marshal.Copy(content, 0, memory, content.Length);
                Marshal.WriteByte(memory, content.Length, 0);

                STGMEDIUM medium = new STGMEDIUM();
                medium.tymed = TYMED.TYMED_HGLOBAL;
                medium.unionmember = memory;
                medium.pUnkForRelease = null;
                dataObject.SetData(ref format, ref medium, false);
            }
            finally
            {
                Marshal.FreeHGlobal(memory);
            }
        }

        private static void CloseOleObject(object oleObject, uint closeOption)
        {
            if (oleObject == null)
            {
                return;
            }
            try
            {
                IOleObject closeable = oleObject as IOleObject;
                if (closeable == null)
                {
                    throw new BridgeException(
                        "ole_close_unavailable",
                        "The MathType object does not expose IOleObject.Close."
                    );
                }
                int hr = closeable.Close(closeOption);
                if (hr < 0)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }
            }
            finally
            {
                ComObjects.Release(oleObject);
            }
        }
    }

    internal static class MathMl
    {
        private const string MathMlNamespace = "http://www.w3.org/1998/Math/MathML";

        public static MathMlValue Parse(string xml)
        {
            if (String.IsNullOrWhiteSpace(xml))
            {
                throw new BridgeException("invalid_mathml", "MathML must not be empty.");
            }

            XmlDocument document = new XmlDocument();
            document.PreserveWhitespace = false;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Prohibit;
            settings.XmlResolver = null;
            try
            {
                using (StringReader text = new StringReader(xml))
                using (XmlReader reader = XmlReader.Create(text, settings))
                {
                    document.Load(reader);
                }
            }
            catch (XmlException error)
            {
                throw new BridgeException("invalid_mathml", "Invalid MathML XML: " + error.Message);
            }

            XmlElement root = document.DocumentElement;
            if (root == null
                || root.LocalName != "math"
                || root.NamespaceURI != MathMlNamespace)
            {
                throw new BridgeException(
                    "invalid_mathml",
                    "MathML root must be {http://www.w3.org/1998/Math/MathML}math."
                );
            }

            XmlDocument normalized = NormalizeNamespaces(document);
            XmlDsigC14NTransform transform = new XmlDsigC14NTransform(false);
            transform.LoadInput(normalized);
            byte[] canonical;
            using (Stream output = (Stream)transform.GetOutput(typeof(Stream)))
            using (MemoryStream buffer = new MemoryStream())
            {
                output.CopyTo(buffer);
                canonical = buffer.ToArray();
            }

            string digest;
            using (SHA256 sha256 = SHA256.Create())
            {
                digest = ToHex(sha256.ComputeHash(canonical));
            }
            return new MathMlValue(
                xml,
                new UTF8Encoding(false).GetString(canonical),
                digest
            );
        }

        private static XmlDocument NormalizeNamespaces(XmlDocument source)
        {
            SortedSet<string> attributeNamespaces = new SortedSet<string>(StringComparer.Ordinal);
            CollectAttributeNamespaces(source.DocumentElement, attributeNamespaces);
            Dictionary<string, string> prefixes = new Dictionary<string, string>();
            int prefixIndex = 1;
            foreach (string namespaceUri in attributeNamespaces)
            {
                prefixes[namespaceUri] = "ns" + prefixIndex;
                prefixIndex++;
            }

            XmlDocument normalized = new XmlDocument();
            normalized.PreserveWhitespace = false;
            normalized.AppendChild(CloneElement(source.DocumentElement, normalized, prefixes));
            return normalized;
        }

        private static void CollectAttributeNamespaces(
            XmlElement element,
            SortedSet<string> namespaces)
        {
            foreach (XmlAttribute attribute in element.Attributes)
            {
                if (!String.IsNullOrEmpty(attribute.NamespaceURI)
                    && attribute.NamespaceURI != "http://www.w3.org/2000/xmlns/"
                    && attribute.NamespaceURI != "http://www.w3.org/XML/1998/namespace")
                {
                    namespaces.Add(attribute.NamespaceURI);
                }
            }
            foreach (XmlNode child in element.ChildNodes)
            {
                XmlElement childElement = child as XmlElement;
                if (childElement != null)
                {
                    CollectAttributeNamespaces(childElement, namespaces);
                }
            }
        }

        private static XmlElement CloneElement(
            XmlElement source,
            XmlDocument targetDocument,
            Dictionary<string, string> attributePrefixes)
        {
            XmlElement target = targetDocument.CreateElement(
                String.Empty,
                source.LocalName,
                source.NamespaceURI
            );
            foreach (XmlAttribute sourceAttribute in source.Attributes)
            {
                if (sourceAttribute.NamespaceURI == "http://www.w3.org/2000/xmlns/")
                {
                    continue;
                }

                string prefix = String.Empty;
                if (sourceAttribute.NamespaceURI == "http://www.w3.org/XML/1998/namespace")
                {
                    prefix = "xml";
                }
                else if (!String.IsNullOrEmpty(sourceAttribute.NamespaceURI))
                {
                    prefix = attributePrefixes[sourceAttribute.NamespaceURI];
                }
                XmlAttribute targetAttribute = targetDocument.CreateAttribute(
                    prefix,
                    sourceAttribute.LocalName,
                    sourceAttribute.NamespaceURI
                );
                targetAttribute.Value = sourceAttribute.Value;
                target.Attributes.Append(targetAttribute);
            }

            foreach (XmlNode sourceChild in source.ChildNodes)
            {
                XmlElement childElement = sourceChild as XmlElement;
                if (childElement != null)
                {
                    target.AppendChild(CloneElement(
                        childElement,
                        targetDocument,
                        attributePrefixes
                    ));
                }
                else if (sourceChild.NodeType == XmlNodeType.Text
                    || sourceChild.NodeType == XmlNodeType.CDATA
                    || sourceChild.NodeType == XmlNodeType.SignificantWhitespace)
                {
                    target.AppendChild(targetDocument.CreateTextNode(sourceChild.Value));
                }
            }
            return target;
        }

        private static string ToHex(byte[] bytes)
        {
            StringBuilder result = new StringBuilder(bytes.Length * 2);
            foreach (byte value in bytes)
            {
                result.Append(value.ToString("x2"));
            }
            return result.ToString();
        }
    }

    internal sealed class MathMlValue
    {
        public readonly string OriginalXml;
        public readonly string CanonicalXml;
        public readonly string Sha256;

        public MathMlValue(string originalXml, string canonicalXml, string sha256)
        {
            OriginalXml = originalXml;
            CanonicalXml = canonicalXml;
            Sha256 = sha256;
        }
    }

    internal sealed class OleMathMl
    {
        public readonly string FormatName;
        public readonly MathMlValue Value;

        public OleMathMl(string formatName, MathMlValue value)
        {
            FormatName = formatName;
            Value = value;
        }
    }

    internal sealed class FormatSelection
    {
        public readonly string Name;
        public readonly FORMATETC Format;

        public FormatSelection(string name, FORMATETC format)
        {
            Name = name;
            Format = format;
        }
    }

    internal sealed class BridgeException : Exception
    {
        public readonly string Code;

        public BridgeException(string code, string message)
            : base(message)
        {
            Code = code;
        }
    }

    [ComImport]
    [Guid("00000112-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleObject
    {
        [PreserveSig]
        int SetClientSite(IntPtr clientSite);

        [PreserveSig]
        int GetClientSite(out IntPtr clientSite);

        [PreserveSig]
        int SetHostNames(
            [MarshalAs(UnmanagedType.LPWStr)] string containerApplication,
            [MarshalAs(UnmanagedType.LPWStr)] string containerObject);

        [PreserveSig]
        int Close(uint saveOption);
    }
}
