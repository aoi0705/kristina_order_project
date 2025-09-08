from django.contrib import admin
from .models import *
# Register your models here.

@admin.register(ImportBatch)
class ImportBatchAdmin(admin.ModelAdmin):
    list_display = ('id', 'created_at', 'source_filename', 'sheet_name', 'note')
    search_fields = ('source_filename', 'note')

@admin.register(OrderContent)
class OrderContentAdmin(admin.ModelAdmin):
    list_display = ('id', 'Jan_code', 'Product_name', 'Brand_name', 'SKU_number', 'batch')
    list_filter = ('batch',)
    search_fields = ('Jan_code', 'Product_name', 'Brand_name', 'SKU_number')