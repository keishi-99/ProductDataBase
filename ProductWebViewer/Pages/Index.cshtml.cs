using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ProductWebViewer.Data;
using ProductWebViewer.Models;

namespace ProductWebViewer.Pages;

public class IndexModel : PageModel {
    private readonly ProductRecordRepository _productRepo;
    private readonly SubstrateRecordRepository _substrateRepo;
    private readonly ILogger<IndexModel> _logger;

    public IndexModel(ProductRecordRepository productRepo, SubstrateRecordRepository substrateRepo, ILogger<IndexModel> logger) {
        _productRepo = productRepo;
        _substrateRepo = substrateRepo;
        _logger = logger;
    }

    [BindProperty(SupportsGet = true)] public string Tab { get; set; } = "product";
    [BindProperty(SupportsGet = true)] public string SubTab { get; set; } = "records";
    [BindProperty(SupportsGet = true)] public string StockGroup { get; set; } = "model";
    // 初回表示（未検索）と「0件ヒット」を区別するためのフラグ
    [BindProperty(SupportsGet = true)] public bool HasSearched { get; set; }
    // リストフィルターの折りたたみ状態（サブタブ切替をまたいで維持）
    [BindProperty(SupportsGet = true)] public bool ListCollapsed { get; set; }

    // ページネーション・ソート
    [BindProperty(SupportsGet = true)] public int PageNum { get; set; } = 1;
    [BindProperty(SupportsGet = true)] public int PageSize { get; set; } = 100;
    [BindProperty(SupportsGet = true)] public string SortCol { get; set; } = "";
    [BindProperty(SupportsGet = true)] public string SortDir { get; set; } = "desc";

    // 製品タブ リストボックス選択値
    [BindProperty(SupportsGet = true)] public string? ListProductCategory { get; set; }
    [BindProperty(SupportsGet = true)] public string? ListProductName { get; set; }
    [BindProperty(SupportsGet = true)] public string? ListProductType { get; set; }

    // 製品タブ テキストフィルター
    [BindProperty(SupportsGet = true)] public string? FilterProductName { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterProductOrderNumber { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterProductNumber { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterProductDateType { get; set; } = "regDate";
    [BindProperty(SupportsGet = true)] public string? FilterProductDateFrom { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterProductDateTo { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterSerial { get; set; }

    // 製品タブ データ
    public IReadOnlyList<string> ProductCategoryList { get; private set; } = [];
    public IReadOnlyList<string> ProductNameList { get; private set; } = [];
    public IReadOnlyList<string> ProductTypeList { get; private set; } = [];
    public IReadOnlyList<ProductRecord> ProductRecords { get; private set; } = [];
    public IReadOnlyList<SerialRecord> SerialRecords { get; private set; } = [];

    // 基板タブ リストボックス選択値
    [BindProperty(SupportsGet = true)] public string? ListSubCategory { get; set; }
    [BindProperty(SupportsGet = true)] public string? ListSubProductName { get; set; }
    [BindProperty(SupportsGet = true)] public string? ListSubstrateName { get; set; }

    // 基板タブ テキストフィルター
    [BindProperty(SupportsGet = true)] public string? FilterSubstrateName { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterSubstrateOrderNumber { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterSubstrateNumber { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterSubstrateDateType { get; set; } = "regDate";
    [BindProperty(SupportsGet = true)] public string? FilterSubstrateDateFrom { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterSubstrateDateTo { get; set; }
    [BindProperty(SupportsGet = true)] public bool ExcludeZeroStock { get; set; }

    // 基板タブ データ
    public IReadOnlyList<string> SubCategoryList { get; private set; } = [];
    public IReadOnlyList<string> SubProductNameList { get; private set; } = [];
    public IReadOnlyList<string> SubstrateNameList { get; private set; } = [];
    public IReadOnlyList<SubstrateRecord> SubstrateRecords { get; private set; } = [];
    public IReadOnlyList<StockRecord> StockRecords { get; private set; } = [];

    public int TotalCount { get; private set; }
    public string? ErrorMessage { get; private set; }

    public IActionResult OnGetExportCsv() {
        if (!HasSearched) return BadRequest("検索条件を指定してください。");
        try {
            byte[] bytes;
            string fileName;

            if (Tab == "substrate") {
                if (SubTab == "stock") {
                    bytes = BuildCsvBytes(BuildStockCsvLines(FetchStockRecords(1, 0), StockGroup != "detail"));
                    fileName = $"在庫数_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                }
                else {
                    bytes = BuildCsvBytes(BuildSubstrateCsvLines(FetchSubstrateRecords(1, 0)));
                    fileName = $"基板登録実績_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                }
            }
            else {
                if (SubTab == "serial") {
                    bytes = BuildCsvBytes(BuildSerialCsvLines(FetchSerialRecords(1, 0)));
                    fileName = $"シリアル履歴_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                }
                else {
                    bytes = BuildCsvBytes(BuildProductCsvLines(FetchProductRecords(1, 0)));
                    fileName = $"製品登録実績_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                }
            }

            return File(bytes, "text/csv", fileName);
        } catch (Exception ex) {
            _logger.LogError(ex, "CSVエクスポート中にエラーが発生しました。");
            return StatusCode(500, "エクスポート中にエラーが発生しました。");
        }
    }

    public IActionResult OnGetUsedSubstrates(long id) {
        try {
            var records = _productRepo.GetUsedSubstrates(id);
            return new JsonResult(records);
        } catch (Exception ex) {
            _logger.LogError(ex, "使用基板データの取得中にエラーが発生しました。");
            return new JsonResult(new { error = "データの取得中にエラーが発生しました。" });
        }
    }

    public void OnGet() {
        try {
            if (Tab == "substrate") {
                LoadSubstrateData();
            }
            else {
                LoadProductData();
            }
        } catch (Exception ex) {
            _logger.LogError(ex, "データの処理中にエラーが発生しました。");
            ErrorMessage = "データの処理中にエラーが発生しました。";
        }
    }

    private void LoadProductData() {
        ProductCategoryList = _productRepo.GetCategoryList();
        ProductNameList = _productRepo.GetProductNameList(ListProductCategory);
        ProductTypeList = _productRepo.GetProductTypeList(ListProductCategory, ListProductName);

        if (!HasSearched) return;

        if (SubTab == "serial") {
            TotalCount = CountSerialRecords();
            ClampPage();
            SerialRecords = FetchSerialRecords(PageNum, PageSize);
        }
        else {
            TotalCount = CountProductRecords();
            ClampPage();
            ProductRecords = FetchProductRecords(PageNum, PageSize);
        }
    }

    private void LoadSubstrateData() {
        SubCategoryList = _substrateRepo.GetCategoryList();
        SubProductNameList = _substrateRepo.GetProductNameList(ListSubCategory);
        SubstrateNameList = _substrateRepo.GetSubstrateNameList(ListSubCategory, ListSubProductName);

        if (!HasSearched) return;

        if (SubTab == "stock") {
            TotalCount = CountStockRecords();
            ClampPage();
            StockRecords = FetchStockRecords(PageNum, PageSize);
        }
        else {
            TotalCount = CountSubstrateRecords();
            ClampPage();
            SubstrateRecords = FetchSubstrateRecords(PageNum, PageSize);
        }
    }

    private int CountProductRecords() =>
        _productRepo.GetCount(
            ListProductCategory, ListProductName, ListProductType,
            FilterProductName, FilterProductOrderNumber, FilterProductNumber,
            FilterProductDateType, FilterProductDateFrom, FilterProductDateTo);

    private IReadOnlyList<ProductRecord> FetchProductRecords(int page, int pageSize) =>
        _productRepo.GetAll(
            ListProductCategory, ListProductName, ListProductType,
            FilterProductName, FilterProductOrderNumber, FilterProductNumber,
            FilterProductDateType, FilterProductDateFrom, FilterProductDateTo,
            SortCol, SortDir, page, pageSize);

    private int CountSerialRecords() =>
        _productRepo.GetSerialCount(
            ListProductCategory, ListProductName, ListProductType,
            FilterProductName, FilterProductOrderNumber, FilterProductNumber,
            FilterProductDateType, FilterProductDateFrom, FilterProductDateTo,
            FilterSerial);

    private IReadOnlyList<SerialRecord> FetchSerialRecords(int page, int pageSize) =>
        _productRepo.GetSerialHistory(
            ListProductCategory, ListProductName, ListProductType,
            FilterProductName, FilterProductOrderNumber, FilterProductNumber,
            FilterProductDateType, FilterProductDateFrom, FilterProductDateTo,
            FilterSerial, SortCol, SortDir, page, pageSize);

    private int CountSubstrateRecords() =>
        _substrateRepo.GetCount(
            ListSubCategory, ListSubProductName, ListSubstrateName,
            FilterSubstrateName, FilterSubstrateOrderNumber, FilterSubstrateNumber,
            FilterSubstrateDateType, FilterSubstrateDateFrom, FilterSubstrateDateTo);

    private IReadOnlyList<SubstrateRecord> FetchSubstrateRecords(int page, int pageSize) =>
        _substrateRepo.GetAll(
            ListSubCategory, ListSubProductName, ListSubstrateName,
            FilterSubstrateName, FilterSubstrateOrderNumber, FilterSubstrateNumber,
            FilterSubstrateDateType, FilterSubstrateDateFrom, FilterSubstrateDateTo,
            SortCol, SortDir, page, pageSize);

    private int CountStockRecords() =>
        _substrateRepo.GetStockCount(
            ListSubCategory, ListSubProductName, ListSubstrateName,
            groupByModel: StockGroup != "detail",
            excludeZeroStock: ExcludeZeroStock,
            substrateName: FilterSubstrateName,
            orderNumber: FilterSubstrateOrderNumber,
            substrateNumber: FilterSubstrateNumber);

    private IReadOnlyList<StockRecord> FetchStockRecords(int page, int pageSize) =>
        _substrateRepo.GetStock(
            ListSubCategory, ListSubProductName, ListSubstrateName,
            groupByModel: StockGroup != "detail",
            excludeZeroStock: ExcludeZeroStock,
            substrateName: FilterSubstrateName,
            orderNumber: FilterSubstrateOrderNumber,
            substrateNumber: FilterSubstrateNumber,
            SortCol, SortDir, page, pageSize);

    // ページネーション番号リストを生成する。
    // null はページ番号の間に表示する "..." (省略記号) を表す。
    // 現在ページ±1 の範囲を常に表示し、先頭・末尾は固定で表示する。
    public IEnumerable<int?> GetPageNumbers() {
        if (PageSize <= 0 || TotalCount <= 0) return [];
        var totalPages = (int)Math.Ceiling((double)TotalCount / PageSize);
        if (totalPages <= 1) return [];
        if (totalPages <= 7) return Enumerable.Range(1, totalPages).Cast<int?>();
        var pages = new List<int?> { 1 };
        if (PageNum > 3) pages.Add(null);
        for (var i = Math.Max(2, PageNum - 1); i <= Math.Min(totalPages - 1, PageNum + 1); i++)
            pages.Add(i);
        if (PageNum < totalPages - 2) pages.Add(null);
        pages.Add(totalPages);
        return pages;
    }

    // BOM付きUTF-8 バイト列を返す（Excelで日本語が文字化けしないようにするため）
    private static byte[] BuildCsvBytes(IEnumerable<string> lines) {
        using var ms = new MemoryStream();
        using (var writer = new StreamWriter(ms, new UTF8Encoding(true))) {
            writer.NewLine = "\r\n";
            foreach (var line in lines)
                writer.WriteLine(line);
        }
        return ms.ToArray();
    }

    // カンマ・ダブルクォート・改行を含むフィールドをエスケープする
    // 数式インジェクション対策: =,+,-,@ 始まりの値にはシングルクォートを付加する
    private static string CsvField(string? value) {
        if (string.IsNullOrEmpty(value)) return "";
        value = value.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ");
        if (value[0] is '=' or '+' or '-' or '@')
            value = "'" + value;
        if (value.Contains(',') || value.Contains('"'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return value;
    }

    private static IEnumerable<string> BuildProductCsvLines(IReadOnlyList<ProductRecord> records) {
        yield return "ID,カテゴリ,製品名,種別,型式,注文番号,製造番号,O-Les番号,数量,担当者,登録日,Revision,シリアル開始,シリアル終了,コメント,登録日時";
        foreach (var r in records)
            yield return string.Join(",",
                CsvField(r.Id.ToString()), CsvField(r.CategoryName), CsvField(r.ProductName),
                CsvField(r.ProductType), CsvField(r.ProductModel), CsvField(r.OrderNumber),
                CsvField(r.ProductNumber), CsvField(r.OLesNumber), CsvField(r.Quantity?.ToString()),
                CsvField(r.PersonInfo), CsvField(r.RegDate), CsvField(r.Revision),
                CsvField(r.SerialFirst), CsvField(r.SerialLast), CsvField(r.Comment), CsvField(r.CreatedAt));
    }

    private static IEnumerable<string> BuildSerialCsvLines(IReadOnlyList<SerialRecord> records) {
        yield return "No.,シリアル番号,O-Lesシリアル,注文番号,製番,製品名,種別,型式,登録日,登録日時";
        foreach (var r in records)
            yield return string.Join(",",
                CsvField(r.RowId.ToString()), CsvField(r.Serial), CsvField(r.OLesSerial),
                CsvField(r.OrderNumber), CsvField(r.ProductNumber), CsvField(r.ProductName),
                CsvField(r.ProductType), CsvField(r.ProductModel), CsvField(r.RegDate), CsvField(r.CreatedAt));
    }

    private static IEnumerable<string> BuildSubstrateCsvLines(IReadOnlyList<SubstrateRecord> records) {
        yield return "ID,カテゴリ,製品名,基板名,基板型式,注文番号,製造番号,入庫,出庫,不良,使用製品名,使用注文番号,使用製造番号,担当者,登録日,コメント,登録日時";
        foreach (var r in records)
            yield return string.Join(",",
                CsvField(r.Id.ToString()), CsvField(r.CategoryName), CsvField(r.ProductName),
                CsvField(r.SubstrateName), CsvField(r.SubstrateModel), CsvField(r.OrderNumber),
                CsvField(r.SubstrateNumber), CsvField(r.Increase?.ToString()), CsvField(r.Decrease?.ToString()),
                CsvField(r.Defect?.ToString()), CsvField(r.UseProductName), CsvField(r.UseOrderNumber),
                CsvField(r.UseProductNumber), CsvField(r.PersonInfo), CsvField(r.RegDate),
                CsvField(r.Comment), CsvField(r.CreatedAt));
    }

    private static IEnumerable<string> BuildStockCsvLines(IReadOnlyList<StockRecord> records, bool groupByModel) {
        yield return groupByModel
            ? "カテゴリ,製品名,基板名,基板型式,在庫数"
            : "カテゴリ,製品名,基板名,基板型式,製造番号,注文番号,在庫数";
        foreach (var r in records) {
            if (groupByModel)
                yield return string.Join(",",
                    CsvField(r.CategoryName), CsvField(r.ProductName),
                    CsvField(r.SubstrateName), CsvField(r.SubstrateModel), CsvField(r.Stock.ToString()));
            else
                yield return string.Join(",",
                    CsvField(r.CategoryName), CsvField(r.ProductName),
                    CsvField(r.SubstrateName), CsvField(r.SubstrateModel),
                    CsvField(r.SubstrateNumber), CsvField(r.OrderNumber), CsvField(r.Stock.ToString()));
        }
    }

    // フィルター変更などで総ページ数が減った場合に PageNum を有効範囲内に収める
    private void ClampPage() {
        if (PageNum < 1 || TotalCount == 0) {
            PageNum = 1;
            return;
        }
        if (PageSize > 0) {
            var totalPages = (int)Math.Ceiling((double)TotalCount / PageSize);
            if (PageNum > totalPages) PageNum = totalPages;
        }
    }
}
