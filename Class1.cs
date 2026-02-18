using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.IO;
using System.Text;

[assembly: CommandClass(typeof(InjectNames.Commands))]

namespace InjectNames
{
    public record BlockData(string OriginalName, string Name, string Length, string Width);

    public class Commands
    {
        private const string TargetBlockName = "Panel";

        [CommandMethod("giveNames", CommandFlags.Modal | CommandFlags.Redraw)]
        public void CreateNamesCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var (foundBlocks, seenNames) = SelectBlocksByName(ed, tr, TargetBlockName);

                if (foundBlocks == null)
                {
                    tr.Commit();
                    return;
                }

                if (foundBlocks.Count == 0)
                {
                    ReportNoBlocksFound(ed, seenNames);
                    tr.Commit();
                    return;
                }

                var lengthMap = new Dictionary<string, int>();
                var widthMap = new Dictionary<string, int>();

                var blockDataList = new List<BlockData>();

                foreach (var br in foundBlocks)
                {
                    string length = RoundToInteger(GetDynamicPropertyValue(br, "Length"));
                    string width = RoundToInteger(GetDynamicPropertyValue(br, "Width"));

                    string lengthCode = GetOrAddLengthCode(lengthMap, length);
                    string widthCode = GetOrAddWidthCode(widthMap, width);

                    string originalName = GetAttributeValue(br, "NAME");
                    string newName = $"{originalName}-{lengthCode}-{widthCode}";

                    SetAttributeValue(tr, br, "NAME", newName);

                    blockDataList.Add(new BlockData(originalName, newName, length, width));
                }

                ed.SetImpliedSelection(foundBlocks.Select(br => br.ObjectId).ToArray());

                string csvPath = GetCsvPath(doc);
                ExportToCsv(csvPath, blockDataList);

                ed.WriteMessage($"\nSuccess: {foundBlocks.Count} '{TargetBlockName}' block(s) processed.");
                ed.WriteMessage($"\nCSV exported to: {csvPath}");

                tr.Commit();
            }
        }

        private static (List<BlockReference>? Found, HashSet<string> SeenNames) SelectBlocksByName(
            Editor ed, Transaction tr, string targetName)
        {
            ed.WriteMessage("\nSelect an area containing the blocks to process:");

            var filter = new SelectionFilter([
                new TypedValue((int)DxfCode.Start, "INSERT")
            ]);

            PromptSelectionResult selResult = ed.GetSelection(filter);

            if (selResult.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nSelection cancelled.");
                return (null, []);
            }

            var found = new List<BlockReference>();
            var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (ObjectId objId in selResult.Value.GetObjectIds())
            {
                if (tr.GetObject(objId, OpenMode.ForRead) is BlockReference br)
                {
                    string effectiveName = GetBlockEffectiveName(tr, br);
                    seenNames.Add(effectiveName);

                    if (effectiveName.Equals(targetName, StringComparison.OrdinalIgnoreCase))
                        found.Add(br);
                }
            }

            return (found, seenNames);
        }

        private static string GetBlockEffectiveName(Transaction tr, BlockReference br)
        {
            ObjectId btrId = br.IsDynamicBlock ? br.DynamicBlockTableRecord : br.BlockTableRecord;
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
            return btr.Name;
        }

        private static string RoundToInteger(string value)
        {
            if (double.TryParse(value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double number))
            {
                return ((long)Math.Round(number)).ToString();
            }
            return value;
        }

        private static string GetOrAddLengthCode(Dictionary<string, int> map, string lengthValue)
        {
            if (!map.TryGetValue(lengthValue, out int code))
            {
                code = map.Count + 1;
                map[lengthValue] = code;
            }
            return code.ToString();
        }

        private static string GetOrAddWidthCode(Dictionary<string, int> map, string widthValue)
        {
            if (!map.TryGetValue(widthValue, out int index))
            {
                index = map.Count;
                map[widthValue] = index;
            }
            return ToExcelColumn(index);
        }

        private static string ToExcelColumn(int index)
        {
            var sb = new StringBuilder();
            int remaining = index;
            do
            {
                sb.Insert(0, (char)('A' + remaining % 26));
                remaining = remaining / 26 - 1;
            } while (remaining >= 0);
            return sb.ToString();
        }

        private static string GetDynamicPropertyValue(BlockReference br, string propertyName)
        {
            if (!br.IsDynamicBlock)
                return "";

            foreach (DynamicBlockReferenceProperty prop in br.DynamicBlockReferencePropertyCollection)
            {
                if (prop.PropertyName.Equals(propertyName, StringComparison.OrdinalIgnoreCase))
                    return prop.Value?.ToString() ?? "";
            }

            return "";
        }

        private static string GetAttributeValue(BlockReference br, string tag)
        {
            foreach (ObjectId attId in br.AttributeCollection)
            {
                if (attId.GetObject(OpenMode.ForRead) is AttributeReference att
                    && att.Tag.Equals(tag, StringComparison.OrdinalIgnoreCase))
                {
                    return att.TextString ?? "";
                }
            }

            return "";
        }

        private static void SetAttributeValue(Transaction tr, BlockReference br, string tag, string value)
        {
            foreach (ObjectId attId in br.AttributeCollection)
            {
                AttributeReference att = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                if (att.Tag.Equals(tag, StringComparison.OrdinalIgnoreCase))
                {
                    att.UpgradeOpen();
                    att.TextString = value;
                    return;
                }
            }
        }

        private static string GetCsvPath(Document doc)
        {
            string dwgPath = doc.Name;
            string directory = Path.GetDirectoryName(dwgPath)!;
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(dwgPath);
            return Path.Combine(directory, $"{fileNameWithoutExt}_Panels.csv");
        }

        private static string GetPrefix(string originalName)
        {
            int lastDash = originalName.LastIndexOf('-');
            if (lastDash > 0)
                return originalName[..lastDash];
            return originalName;
        }

        private static double CalculateArea(string length, string width, int quantity)
        {
            if (double.TryParse(length, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out double l)
                && double.TryParse(width, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out double w))
            {
                return l * w * quantity / 1_000_000.0;
            }
            return 0;
        }

        private static void ExportToCsv(string path, List<BlockData> data)
        {
            var groups = data
                .GroupBy(b => GetPrefix(b.OriginalName))
                .OrderBy(g => g.Key);

            var sb = new StringBuilder();
            bool first = true;
            double grandTotal = 0;

            foreach (var group in groups)
            {
                if (!first)
                    sb.AppendLine();
                first = false;

                sb.AppendLine($"{EscapeCsv(group.Key)}");
                sb.AppendLine("NAME,Length,Width,Quantity,Area");

                var merged = group
                    .GroupBy(b => b.Name)
                    .Select(g => (Block: g.First(), Quantity: g.Count()));

                double groupTotal = 0;

                foreach (var (block, quantity) in merged)
                {
                    double area = CalculateArea(block.Length, block.Width, quantity);
                    groupTotal += area;
                    sb.AppendLine($"{EscapeCsv(block.Name)},{EscapeCsv(block.Length)},{EscapeCsv(block.Width)},{quantity},{area:F2}");
                }

                sb.AppendLine($",,,,{groupTotal:F2}");
                grandTotal += groupTotal;
            }
            sb.AppendLine();
            sb.AppendLine($"Total,,,,{grandTotal:F2}");

            File.WriteAllText(path, sb.ToString());
        }

        private static string EscapeCsv(string value)
        {
            if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
                return $"\"{value.Replace("\"", "\"\"")}\"";
            return value;
        }

        private static void ReportNoBlocksFound(Editor ed, HashSet<string> seenNames)
        {
            ed.WriteMessage($"\nNo blocks named '{TargetBlockName}' found in Model Space.");
            if (seenNames.Count > 0)
                ed.WriteMessage("\nBlock names found in drawing: " + string.Join(", ", seenNames));
            else
                ed.WriteMessage("\nNo block references found at all in Model Space.");
        }
    }
}