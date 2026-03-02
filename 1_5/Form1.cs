using System.Text;
using System.Xml.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace _1_5;

public partial class Form1 : Form
{
    // Пространства имён из XML
    private static readonly XNamespace NsBody = "http://v8.1c.ru/edi/edi_stnd/EnterpriseData/1.17";
    private static readonly XNamespace NsMsg = "http://www.1c.ru/SSL/Exchange/Message";

    private string? _xmlPath;
    private XDocument? _xmlDoc;

    // Словарь замены: Артикул -> новая Ссылка номенклатуры
    private Dictionary<string, string> _refMap = new(StringComparer.Ordinal);

    private readonly Button _btnOpenXml = new()
        { Text = "📂 Открыть XML", Width = 170, Height = 36, Left = 12, Top = 12 };

    private readonly Button _btnLoadRef = new()
        { Text = "🔄 Загрузить замены", Width = 170, Height = 36, Left = 192, Top = 12 };

    private readonly Button _btnExport = new()
        { Text = "💾 Сохранить XML", Width = 170, Height = 36, Left = 372, Top = 12, Enabled = false };

    private readonly Label _lblXml = new()
        { Text = "XML файл: не выбран", Left = 12, Top = 56, Width = 760, Height = 20, AutoSize = false };

    private readonly Label _lblRef = new()
    {
        Text = "Файл замен: не загружен", Left = 12, Top = 78, Width = 760, Height = 20, AutoSize = false,
        ForeColor = System.Drawing.Color.Gray
    };

    private readonly Label _lblStatus = new()
    {
        Left = 12, Top = 100, Width = 760, Height = 20, AutoSize = false, ForeColor = System.Drawing.Color.DarkGreen
    };

    private readonly ListBox _lstDocs = new()
        { Left = 12, Top = 128, Width = 760, Height = 400, Font = new System.Drawing.Font("Consolas", 9f) };


    public Form1()
    {
        Text = "XML → РеализацияТоваровУслуг (замена ссылок)";
        ClientSize = new System.Drawing.Size(800, 560);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        Controls.AddRange(
            new Control[] { _btnOpenXml, _btnLoadRef, _btnExport, _lblXml, _lblRef, _lblStatus, _lstDocs });
        _btnOpenXml.Click += BtnOpenXml_Click;
        _btnLoadRef.Click += BtnLoadRef_Click;
        _btnExport.Click += BtnExport_Click;
    }

    // ─── Открыть XML ──────────────────────────────────────────────────────
    private void BtnOpenXml_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title = "Выберите XML файл",
            Filter = "XML файлы (*.xml)|*.xml|Все файлы (*.*)|*.*"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        _xmlPath = dlg.FileName;
        _lblXml.Text = "XML файл: " + _xmlPath;
        SetStatus("Загрузка...", false);
        _lstDocs.Items.Clear();

        try
        {
            _xmlDoc = XDocument.Load(_xmlPath);

            var docs = GetRealizaciyaDocs(_xmlDoc);

            _lstDocs.Items.Add($"Всего блоков Документ.РеализацияТоваровУслуг: {docs.Count}");
            _lstDocs.Items.Add(new string('-', 90));
            foreach (var doc in docs)
            {
                string num = doc.Descendants(NsBody + "Номер").FirstOrDefault()?.Value ?? "?";
                string date = doc.Descendants(NsBody + "Дата").FirstOrDefault()?.Value ?? "?";
                int rows = doc.Descendants(NsBody + "Строка").Count();
                _lstDocs.Items.Add($"№ {num}  |  {date}  |  товаров: {rows}");
            }

            _btnExport.Enabled = true;
            SetStatus($"✔ Найдено {docs.Count} документов реализации", false);
        }
        catch (Exception ex)
        {
            SetStatus("Ошибка загрузки XML: " + ex.Message, true);
        }
    }

    // ─── Загрузить файл замен (необязательно) ─────────────────────────────
    private void BtnLoadRef_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title = "Выберите Excel файл с заменами ссылок",
            Filter = "Excel файлы (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы (*.*)|*.*"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            _refMap = ReadRefMap(dlg.FileName);
            _lblRef.Text = $"Файл замен: {dlg.FileName}  ({_refMap.Count} записей)";
            _lblRef.ForeColor = System.Drawing.Color.DarkGreen;
            SetStatus($"✔ Загружено {_refMap.Count} замен артикулов", false);
        }
        catch (Exception ex)
        {
            SetStatus("Ошибка загрузки замен: " + ex.Message, true);
        }
    }

    // ─── Сохранить XML ────────────────────────────────────────────────────
    private void BtnExport_Click(object? sender, EventArgs e)
    {
        if (_xmlDoc == null) return;

        using var dlg = new SaveFileDialog
        {
            Title = "Сохранить XML",
            Filter = "XML файлы (*.xml)|*.xml",
            FileName = $"ГОТОВО {Path.GetFileNameWithoutExtension(_xmlPath)}.xml",
            DefaultExt = "xml"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            string result = BuildXml(_xmlDoc, _refMap);
            File.WriteAllText(dlg.FileName, result, new UTF8Encoding(false));
            SetStatus($"✔ Сохранено: {dlg.FileName}", false);
            MessageBox.Show($"Готово!\n{dlg.FileName}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            SetStatus("Ошибка: " + ex.Message, true);
            MessageBox.Show(ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ─── Получить все блоки Документ.РеализацияТоваровУслуг ──────────────
    private static List<XElement> GetRealizaciyaDocs(XDocument doc)
        => doc.Descendants(NsBody + "Документ.РеализацияТоваровУслуг").ToList();

    // ─── Собрать итоговый XML: только РеализацияТоваровУслуг + замена ссылок
    private static string BuildXml(XDocument sourceDoc, Dictionary<string, string> refMap)
    {
        // Берём оригинальный корень и Header как есть
        var root = sourceDoc.Root ?? throw new Exception("Пустой XML");
        var header = root.Element(NsMsg + "Header");
        var body = root.Element(NsBody + "Body");

        if (body == null) throw new Exception("Элемент Body не найден");

        // Оставляем только нужные документы, клонируем их
        var docs = body.Elements(NsBody + "Документ.РеализацияТоваровУслуг")
            .Select(d => ApplyRefMap(new XElement(d), refMap)) // глубокий клон + замена
            .ToList();

        // Строим новый XML с той же структурой
        var newRoot = new XElement(root.Name,
            root.Attributes(), // xmlns:msg, xmlns:xs, xmlns:xsi
            header != null ? new XElement(header) : null, // Header как есть
            new XElement(NsBody + "Body",
                body.Attributes(),
                docs
            )
        );

        var sb = new StringBuilder();
        var settings = new System.Xml.XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false
        };

        using var xw = System.Xml.XmlWriter.Create(sb, settings);
        new XDocument(
            new XDeclaration("1.0", "utf-8", null),
            newRoot
        ).WriteTo(xw);
        xw.Flush();

        return sb.ToString();
    }

    // ─── Применить замену ссылок к одному документу ───────────────────────
    // Ищем все Номенклатура внутри документа,
    // берём Артикул, ищем в словаре, если нашли — меняем Ссылка
    private static XElement ApplyRefMap(XElement doc, Dictionary<string, string> refMap)
    {
        if (refMap.Count == 0) return doc; // замен нет — возвращаем как есть

        foreach (var nom in doc.Descendants(NsBody + "Номенклатура"))
        {
            var artikulEl = nom.Element(NsBody + "Артикул");
            var ssylkaEl = nom.Element(NsBody + "Ссылка");

            if (artikulEl == null || ssylkaEl == null) continue;

            string article = artikulEl.Value.Trim();
            if (refMap.TryGetValue(article, out string? newRef))
                ssylkaEl.Value = newRef;
        }

        return doc;
    }

    // ─── Чтение файла замен: Артикул (col 4) -> Ссылка (col 2) ───────────
    private static Dictionary<string, string> ReadRefMap(string path)
    {
        IWorkbook workbook;
        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            workbook = path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? new XSSFWorkbook(fs)
                : new HSSFWorkbook(fs);
        }

        var sheet = workbook.GetSheetAt(1);
        var formatter = new DataFormatter();
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        int lastRow = sheet.LastRowNum;

        for (int r = 1; r <= lastRow; r++) // строка 0 — заголовок
        {
            var row = sheet.GetRow(r);
            if (row == null) continue;

            string newRef = formatter.FormatCellValue(row.GetCell(2))?.Trim() ?? string.Empty; // col C
            string article = formatter.FormatCellValue(row.GetCell(4))?.Trim() ?? string.Empty; // col E

            if (!string.IsNullOrEmpty(article) && !string.IsNullOrEmpty(newRef))
                map[article] = newRef;
        }

        return map;
    }

    private void SetStatus(string text, bool isError)
    {
        _lblStatus.Text = text;
        _lblStatus.ForeColor = isError ? System.Drawing.Color.Red : System.Drawing.Color.DarkGreen;
    }
}