using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ProductWebViewer.Data;
using ProductWebViewer.Models;

namespace ProductWebViewer.Pages;

public class IndexModel : PageModel {
    private readonly ProductRecordRepository _productRepo;
    private readonly SubstrateRecordRepository _substrateRepo;

    public IndexModel(ProductRecordRepository productRepo, SubstrateRecordRepository substrateRepo) {
        _productRepo = productRepo;
        _substrateRepo = substrateRepo;
    }

    [BindProperty(SupportsGet = true)] public string Tab { get; set; } = "product";
    [BindProperty(SupportsGet = true)] public string SubTab { get; set; } = "records";
    [BindProperty(SupportsGet = true)] public string StockGroup { get; set; } = "model";

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

    public string? ErrorMessage { get; private set; }

    public void OnGet() {
        try {
            if (Tab == "substrate") {
                SubCategoryList = _substrateRepo.GetCategoryList();
                SubProductNameList = _substrateRepo.GetProductNameList(ListSubCategory);
                SubstrateNameList = _substrateRepo.GetSubstrateNameList(ListSubCategory, ListSubProductName);

                if (SubTab == "stock") {
                    StockRecords = _substrateRepo.GetStock(
                        ListSubCategory, ListSubProductName, ListSubstrateName,
                        groupByModel: StockGroup != "detail");
                } else {
                    SubstrateRecords = _substrateRepo.GetAll(
                        ListSubCategory, ListSubProductName, ListSubstrateName,
                        FilterSubstrateName, FilterSubstrateOrderNumber,
                        FilterSubstrateRegDateFrom, FilterSubstrateRegDateTo);
                }
            } else {
                ProductCategoryList = _productRepo.GetCategoryList();
                ProductNameList = _productRepo.GetProductNameList(ListProductCategory);
                ProductTypeList = _productRepo.GetProductTypeList(ListProductCategory, ListProductName);

                if (SubTab == "serial") {
                    SerialRecords = _productRepo.GetSerialHistory(
                        ListProductName, ListProductType, FilterSerial);
                } else {
                    ProductRecords = _productRepo.GetAll(
                        ListProductCategory, ListProductName, ListProductType,
                        FilterProductName, FilterProductOrderNumber,
                        FilterProductRegDateFrom, FilterProductRegDateTo);
                }
            }
        } catch (Exception ex) {
            ErrorMessage = ex.Message;
        }
    }
}
