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

    // ページネーション・ソート
    [BindProperty(SupportsGet = true)] public int    PageNum  { get; set; } = 1;
    [BindProperty(SupportsGet = true)] public int    PageSize { get; set; } = 100;
    [BindProperty(SupportsGet = true)] public string SortCol  { get; set; } = "";
    [BindProperty(SupportsGet = true)] public string SortDir  { get; set; } = "desc";

    // 製品タブ リストボックス選択値
    [BindProperty(SupportsGet = true)] public string? ListProductCategory { get; set; }
    [BindProperty(SupportsGet = true)] public string? ListProductName { get; set; }
    [BindProperty(SupportsGet = true)] public string? ListProductType { get; set; }

    // 製品タブ テキストフィルター
    [BindProperty(SupportsGet = true)] public string? FilterProductName { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterProductOrderNumber { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterProductRegDateFrom { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterProductRegDateTo { get; set; }
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
    [BindProperty(SupportsGet = true)] public string? FilterSubstrateRegDateFrom { get; set; }
    [BindProperty(SupportsGet = true)] public string? FilterSubstrateRegDateTo { get; set; }

    // 基板タブ データ
    public IReadOnlyList<string> SubCategoryList { get; private set; } = [];
    public IReadOnlyList<string> SubProductNameList { get; private set; } = [];
    public IReadOnlyList<string> SubstrateNameList { get; private set; } = [];
    public IReadOnlyList<SubstrateRecord> SubstrateRecords { get; private set; } = [];
    public IReadOnlyList<StockRecord> StockRecords { get; private set; } = [];

    public int TotalCount { get; private set; }
    public string? ErrorMessage { get; private set; }

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
            } else {
                LoadProductData();
            }
        } catch (Exception ex) {
            _logger.LogError(ex, "データの処理中にエラーが発生しました。");
            ErrorMessage = "データの処理中にエラーが発生しました。";
        }
    }

    private void LoadProductData() {
        ProductCategoryList = _productRepo.GetCategoryList();
        ProductNameList     = _productRepo.GetProductNameList(ListProductCategory);
        ProductTypeList     = _productRepo.GetProductTypeList(ListProductCategory, ListProductName);

        if (!HasSearched) return;

        if (SubTab == "serial") {
            TotalCount = _productRepo.GetSerialCount(
                ListProductName, ListProductType, FilterSerial);
            ClampPage();
            SerialRecords = _productRepo.GetSerialHistory(
                ListProductName, ListProductType, FilterSerial,
                SortCol, SortDir, PageNum, PageSize);
        } else {
            TotalCount = _productRepo.GetCount(
                ListProductCategory, ListProductName, ListProductType,
                FilterProductName, FilterProductOrderNumber,
                FilterProductRegDateFrom, FilterProductRegDateTo);
            ClampPage();
            ProductRecords = _productRepo.GetAll(
                ListProductCategory, ListProductName, ListProductType,
                FilterProductName, FilterProductOrderNumber,
                FilterProductRegDateFrom, FilterProductRegDateTo,
                SortCol, SortDir, PageNum, PageSize);
        }
    }

    private void LoadSubstrateData() {
        SubCategoryList    = _substrateRepo.GetCategoryList();
        SubProductNameList = _substrateRepo.GetProductNameList(ListSubCategory);
        SubstrateNameList  = _substrateRepo.GetSubstrateNameList(ListSubCategory, ListSubProductName);

        if (!HasSearched) return;

        if (SubTab == "stock") {
            TotalCount = _substrateRepo.GetStockCount(
                ListSubCategory, ListSubProductName, ListSubstrateName,
                groupByModel: StockGroup != "detail");
            ClampPage();
            StockRecords = _substrateRepo.GetStock(
                ListSubCategory, ListSubProductName, ListSubstrateName,
                groupByModel: StockGroup != "detail",
                SortCol, SortDir, PageNum, PageSize);
        } else {
            TotalCount = _substrateRepo.GetCount(
                ListSubCategory, ListSubProductName, ListSubstrateName,
                FilterSubstrateName, FilterSubstrateOrderNumber,
                FilterSubstrateRegDateFrom, FilterSubstrateRegDateTo);
            ClampPage();
            SubstrateRecords = _substrateRepo.GetAll(
                ListSubCategory, ListSubProductName, ListSubstrateName,
                FilterSubstrateName, FilterSubstrateOrderNumber,
                FilterSubstrateRegDateFrom, FilterSubstrateRegDateTo,
                SortCol, SortDir, PageNum, PageSize);
        }
    }

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
