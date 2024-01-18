using System.Reflection;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Text;
using Microsoft.CSharp;
using ExcelDataReader;

namespace ExcelJson
{
    public class ExcelCodeParser : ExcelJsonReader
    {
        readonly ExcelCodeOptions m_CodeOptions;

        public ExcelCodeParser() : this(new(), new()) { }
        public ExcelCodeParser(ExcelCodeOptions codeOptions) : this(codeOptions, new()) { }
        public ExcelCodeParser(ExcelJsonOptions jsonOptions) : this(new(), jsonOptions) { }
        public ExcelCodeParser(ExcelCodeOptions codeOptions, ExcelJsonOptions jsonOptions) : base(jsonOptions)
        {
            m_CodeOptions = codeOptions;
        }

        public IEnumerable<HeaderRow> ParseExcel(Stream stream)
        {
            using var reader = ExcelReaderFactory.CreateReader(stream);
            reader.Reset();
            return ReadHeaderRows(reader);
        }

        public void WriteCode(HeaderRow header, string directory)
        {
            var code = CreateCode(header);
            var options = new CodeGeneratorOptions
            {
                BracingStyle = "C",
                BlankLinesBetweenMembers = false,
            };
            var path = Path.Combine(directory, header.Name + ".cs");
            using var file = File.Create(path);
            using var writer = new StreamWriter(file);
            new CSharpCodeProvider().GenerateCodeFromCompileUnit(code, writer, options);
        }

        public void WriteManagerCode(IEnumerable<HeaderRow> headers, string directory)
        {
            var managerCode = CreateManagerHeaderRow(headers);
            WriteCode(managerCode, directory);
        }

        IEnumerable<HeaderRow> ReadHeaderRows(IExcelDataReader reader)
        {
            do
            {
                var sheetName = reader.Name;
                if (sheetName.StartsWith(m_Options.FilterSheetToken))
                {
                    continue;
                }
                var definitions = ReadDefinitions(reader);
                var headerRow = ReadHeaderRow(sheetName, definitions);
                yield return headerRow;
            } while (reader.NextResult());
        }

        HeaderRow CreateManagerHeaderRow(IEnumerable<HeaderRow> headers)
        {
            var fields = new List<HeaderField>();
            var sb = new StringBuilder(1024);

            foreach (var header in headers)
            {
                var collectionType = sb.Append("Dictionary<")
                    .Append(header.Fields[0].Type)
                    .Append(header.Name)
                    .Append(">")
                    .ToString();

                fields.Add(new(collectionType, header.Name, 1));
            }
            return new(m_CodeOptions.ManagerName, fields.ToArray());
        }

        CodeCompileUnit CreateCode(HeaderRow header)
        {
            var code = new CodeCompileUnit();
            var ns = new CodeNamespace();
            ns.Imports.Add(new("System"));
            ns.Imports.Add(new("System.Collection.Generic"));
            ns.Imports.Add(new("Newtonsoft.Json"));
            foreach (var nsImport in m_CodeOptions.NamespaceImports)
            {
                ns.Imports.Add(new(nsImport));
            }
            code.Namespaces.Add(ns);

            var classCode = CreateClassCode(header.Name);
            foreach (var field in header.Fields)
            {
                var memberCode = CreateMemberField(field);
                classCode.Members.Add(memberCode);
            }
            return code;
        }

        CodeTypeDeclaration CreateClassCode(string name)
        {
            return new CodeTypeDeclaration
            {
                Name = name,
                IsClass = true,
                TypeAttributes = TypeAttributes.Public,
                CustomAttributes = new CodeAttributeDeclarationCollection
                {
                    new("Serializable")
                }
            };
        }

        CodeMemberField CreateMemberField(HeaderField field)
        {
            return new CodeMemberField(field.Type, field.Name)
            {
                Attributes = MemberAttributes.Public,
            };
        }
    }
}