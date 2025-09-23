from django.db import models

# Create your models here.
class OrderContent(models.Model):
    Jan_code = models.CharField(max_length=30,null=True)
    Product_number = models.CharField(max_length=30,null=True)
    Brand_name = models.CharField(max_length=256,null=True)
    SKU_number = models.CharField(max_length=256,null=True)
    Product_name = models.CharField(max_length=256,null=True)
    日本語名 = models.CharField(max_length=256,null=True)
    Order = models.CharField(max_length=256,null=True)
    English_name = models.CharField(max_length=256,null=True)
    Rassian_name = models.CharField(max_length=256,null=True)
    Contents = models.CharField(max_length=256,null=True)
    Volume = models.CharField(max_length=256,null=True)
    Case_qty = models.CharField(max_length=256,null=True)
    Lot = models.CharField(max_length=256,null=True)
    Oeder = models.CharField(max_length=256,null=True)
    Unit_price = models.CharField(max_length=256,null=True)
    Amount = models.CharField(max_length=256,null=True)

    Case_volume = models.CharField(max_length=256,null=True)
    Case_weight = models.CharField(max_length=256,null=True)
    Case_qty2 = models.CharField(max_length=256,null=True)
    TTL_volume = models.CharField(max_length=256,null=True)
    TTL_weight = models.CharField(max_length=256,null=True)
    Product_size = models.CharField(max_length=256,null=True)
    Unit_nw = models.CharField(max_length=256,null=True)
    TTL_nw = models.CharField(max_length=256,null=True)
    Ingredients = models.TextField(null=True,blank=True)

    #仕入れ関係追加列
    Purchase_price = models.CharField(max_length=256,null=True)
    Purchase_amount = models.CharField(max_length=256,null=True)
    profit = models.CharField(max_length=256,null=True)
    profit_rate = models.CharField(max_length=256,null=True)

    #以下ロシア用
    DS_details = models.CharField(max_length=256,null=True)
    DS_brandname = models.CharField(max_length=256,null=True)
    DS_Manufacturer = models.CharField(max_length=256,null=True) 

    #商品登録が必要な商品
    RequiredProduct_flg = models.BooleanField(default=False)

    #以下作成日時・更新日時
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True,null=True)

    # ImportBatch への外部キー
    batch = models.ForeignKey('ImportBatch', null=True, blank=True, on_delete=models.SET_NULL, related_name='items')

class NIPPONIKATRADING_OrderContent(models.Model):
    HS_CODE = models.CharField(max_length=256,null=True)
    Jan_code = models.CharField(max_length=30,null=True)
    Артикул = models.CharField(max_length=256,null=True)
    Brand_name = models.CharField(max_length=256,null=True)
    日本語名 = models.CharField(max_length=256,null=True)
    Description_of_goods = models.CharField(max_length=256,null=True)
    Наименование_ДС_англ = models.CharField(max_length=256,null=True)
    Наименование_ДС_рус = models.CharField(max_length=256,null=True)
    Contents = models.CharField(max_length=256,null=True)
    LOT = models.CharField(max_length=256,null=True)
    Case_Qty = models.CharField(max_length=256,null=True)
    ORDER = models.CharField(max_length=256,null=True)
    Unit_price = models.CharField(max_length=256,null=True)
    Amount = models.CharField(max_length=256,null=True)
    仕入値 = models.CharField(max_length=256,null=True)
    仕入値合計 = models.CharField(max_length=256,null=True)
    利益 = models.CharField(max_length=256,null=True)
    利益率 = models.CharField(max_length=256,null=True)
    ケース容積 = models.CharField(max_length=256,null=True)
    ケース重量 = models.CharField(max_length=256,null=True)
    ケース数量 = models.CharField(max_length=256,null=True)
    合計容積 = models.CharField(max_length=256,null=True)
    合計重量 = models.CharField(max_length=256,null=True)
    商品サイズ = models.CharField(max_length=256,null=True)
    Unit_NW = models.CharField(max_length=256,null=True)
    Total_NW = models.CharField(max_length=256,null=True)
    成分 = models.TextField(null=True,blank=True)
    Марка_бренд_ДС = models.CharField(max_length=256,null=True)
    Производель_ДС = models.CharField(max_length=256,null=True)

    #商品登録が必要な商品
    RequiredProduct_flg = models.BooleanField(default=False)

    #以下作成日時・更新日時
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True,null=True)

    # ImportBatch への外部キー
    batch = models.ForeignKey('ImportBatch', null=True, blank=True, on_delete=models.SET_NULL, related_name='nipponika_items')

class YAMATO_TOYO_OrderContent(models.Model):
    Brand = models.CharField(max_length=256,null=True)
    Order_Code = models.CharField(max_length=256,null=True)
    Item_Name = models.CharField(max_length=256,null=True)
    Quantity = models.CharField(max_length=256,null=True)
    Unit_price_JPY = models.CharField(max_length=256,null=True)
    Amount_JPY = models.CharField(max_length=256,null=True)
    販売価格 = models.CharField(max_length=256,null=True)
    総販売価格 = models.CharField(max_length=256,null=True)
    輸出額 = models.CharField(max_length=256,null=True)
    利益 = models.CharField(max_length=256,null=True)
    利益率 = models.CharField(max_length=256,null=True)
    pcs_ct = models.CharField(max_length=256,null=True)
    CTN = models.CharField(max_length=256,null=True)

    #商品登録が必要な商品
    RequiredProduct_flg = models.BooleanField(default=False)

    #以下作成日時・更新日時
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True,null=True)

    # ImportBatch への外部キー
    batch = models.ForeignKey('ImportBatch', null=True, blank=True, on_delete=models.SET_NULL, related_name='yamato_toyo_items')

class ImportBatch(models.Model):
    created_at = models.DateTimeField(auto_now_add=True, db_index=True)
    source_filename = models.CharField(max_length=255, blank=True)
    sheet_name = models.CharField(max_length=128, blank=True)
    note = models.CharField(max_length=255, blank=True)

    buyers = models.CharField(max_length=255, blank=True)

    PurchaseOrder_file = models.CharField(max_length=255, blank=True)
    InvoicePacking_file = models.CharField(max_length=255, blank=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        base = self.source_filename or "(no filename)"
        return f"{self.created_at:%Y-%m-%d %H:%M} - {base}"
    
#以下商品情報管理
class NIPPONIKATRADING_ProductInfo(models.Model):
    HS_CODE = models.CharField(max_length=256,null=True)
    Jan_code = models.CharField(max_length=30,null=True)
    Артикул = models.CharField(max_length=256,null=True)
    Brand_name = models.CharField(max_length=256,null=True)
    日本語名 = models.CharField(max_length=256,null=True)
    Description_of_goods = models.CharField(max_length=256,null=True)
    Наименование_ДС_англ = models.CharField(max_length=256,null=True)
    Наименование_ДС_рус = models.CharField(max_length=256,null=True)
    Contents = models.CharField(max_length=256,null=True)
    LOT = models.CharField(max_length=256,null=True)
    Case_Qty = models.CharField(max_length=256,null=True)
    ORDER = models.CharField(max_length=256,null=True)
    Unit_price = models.CharField(max_length=256,null=True)
    Amount = models.CharField(max_length=256,null=True)
    仕入値 = models.CharField(max_length=256,null=True)
    仕入値合計 = models.CharField(max_length=256,null=True)
    利益 = models.CharField(max_length=256,null=True)
    利益率 = models.CharField(max_length=256,null=True)
    ケース容積 = models.CharField(max_length=256,null=True)
    ケース重量 = models.CharField(max_length=256,null=True)
    ケース数量 = models.CharField(max_length=256,null=True)
    合計容積 = models.CharField(max_length=256,null=True)
    合計重量 = models.CharField(max_length=256,null=True)
    商品サイズ = models.CharField(max_length=256,null=True)
    Unit_NW = models.CharField(max_length=256,null=True)
    Total_NW = models.CharField(max_length=256,null=True)
    成分 = models.TextField(null=True,blank=True)
    Марка_бренд_ДС = models.CharField(max_length=256,null=True)
    Производель_ДС = models.CharField(max_length=256,null=True)

    RequiredProduct_flg = models.BooleanField(default=False)

    #以下作成日時・更新日時
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True,null=True)

class RY_ProductInfo(models.Model):
    Jan_code = models.CharField(max_length=30,null=True)
    Product_number = models.CharField(max_length=30,null=True)
    Brand_name = models.CharField(max_length=256,null=True)
    SKU_number = models.CharField(max_length=256,null=True)
    Product_name = models.CharField(max_length=256,null=True)
    日本語名 = models.CharField(max_length=256,null=True)
    ORDER = models.CharField(max_length=256,null=True)
    English_name = models.CharField(max_length=256,null=True)
    Rassian_name = models.CharField(max_length=256,null=True)
    Contents = models.CharField(max_length=256,null=True)
    Volume = models.CharField(max_length=256,null=True)
    Case_qty = models.CharField(max_length=256,null=True)
    Lot = models.CharField(max_length=256,null=True)
    Unit_price = models.CharField(max_length=256,null=True)
    Amount = models.CharField(max_length=256,null=True)
    Case_volume = models.CharField(max_length=256,null=True)
    Case_weight = models.CharField(max_length=256,null=True)
    Case_qty2 = models.CharField(max_length=256,null=True)
    TTL_volume = models.CharField(max_length=256,null=True)
    TTL_weight = models.CharField(max_length=256,null=True)
    Product_size = models.CharField(max_length=256,null=True)
    Unit_nw = models.CharField(max_length=256,null=True)
    TTL_nw = models.CharField(max_length=256,null=True)
    Ingredients = models.TextField(null=True,blank=True)
    Purchase_price = models.CharField(max_length=256,null=True)
    Purchase_amount = models.CharField(max_length=256,null=True)
    profit = models.CharField(max_length=256,null=True)
    profit_rate = models.CharField(max_length=256,null=True)
    DS_details = models.CharField(max_length=256,null=True)
    DS_brandname = models.CharField(max_length=256,null=True)
    DS_Manufacturer = models.CharField(max_length=256,null=True)

    RequiredProduct_flg = models.BooleanField(default=False)

    #以下作成日時・更新日時
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True,null=True)

class YAMATO_TOYO_ProductInfo(models.Model):
    Brand = models.CharField(max_length=256,null=True)
    Order_Code = models.CharField(max_length=256,null=True)
    Item_Name = models.CharField(max_length=256,null=True)
    Quantity = models.CharField(max_length=256,null=True)
    Unit_price_JPY = models.CharField(max_length=256,null=True)
    Amount_JPY = models.CharField(max_length=256,null=True)
    販売価格 = models.CharField(max_length=256,null=True)
    輸出額 = models.CharField(max_length=256,null=True)
    利益 = models.CharField(max_length=256,null=True)
    利益率 = models.CharField(max_length=256,null=True)
    pcs_ct = models.CharField(max_length=256,null=True)
    CTN = models.CharField(max_length=256,null=True)

    RequiredProduct_flg = models.BooleanField(default=False)

    #以下作成日時・更新日時
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True,null=True)
