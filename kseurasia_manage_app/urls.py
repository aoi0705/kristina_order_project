from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('orders/import/', views.import_orders, name='import_orders'),
    # 一覧をデフォルト画面に
    path("orders/", views.order_list, name="order_list"),
    # 詳細
    path("orders/<int:pk>/", views.order_detail, name="order_detail"),

    #インポートバッチ関連の処理
    path("imports/", views.import_batch_list, name="import_batch_list"),
    # ↓ 生成系のURLはあなたがバックエンド実装（ビューはまだ不要なら 404 でOK）
    path("imports/<int:batch_id>/download/purchase-order/", views.download_purchase_order, name="download_purchase_order"),
    path("imports/<int:batch_id>/export/invoice-packing/", views.export_invoice_packing, name="export_invoice_packing"),

    path("imports/<int:batch_id>/orders/", views.batch_order_list, name="batch_order_list"),
    path("imports/<int:batch_id>/orders/<int:pk>/", views.batch_order_detail, name="batch_order_detail"),

    path('imports/<int:batch_id>/delete/', views.delete_import_batch, name='delete_import_batch'),
    
    #売上表関連
    path("reports/", views.report_console, name="report_console"),
    # --- レポート：ダウンロード（CSV/PDF） ---
    path("api/reports/sales/export",    views.reports_sales_export,    name="reports_sales_export"),
    path("api/reports/ar/export",       views.reports_ar_export,       name="reports_ar_export"),
    path("api/reports/ap/export",       views.reports_ap_export,       name="reports_ap_export"),
    path("api/reports/cashflow/export", views.reports_cashflow_export, name="reports_cashflow_export"),
    #4レポート全てダウンロード
    path("api/reports/export-all", views.reports_export_bundle, name="reports_export_bundle"),

    #以下商品情報管理
    path("products/import/", views.product_import, name="product_import"),
    path("products/", views.product_list, name="product_list"),
    path("products/<int:pk>/", views.product_detail, name="product_detail"),

    #ランキング
    path("rankings/", views.ranking_console, name="ranking_console"),
    path("api/rankings/export", views.rankings_export, name="rankings_export"),
]