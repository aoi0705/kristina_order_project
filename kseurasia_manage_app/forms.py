from django import forms

class OrderExcelUploadForm(forms.Form):

    # DESTINATION_CHOICES = (
    #     ("russia", "ロシア向け"),
    #     ("dubai",  "ドバイ向け"),
    # )

    # CLIENT_CHOICES = (
    #     ("NIPPONIKA TRADING社", "NIPPONIKA TRADING社"),
    #     ("ROYAL COOSMETICS社",  "ROYAL COOSMETICS社"),
    #     ("個人",                 "個人"),
    #     ("USA",                  "USA"),
    #     ("ACES Beteiligunen UG社","ACES Beteiligunen UG社"),
    #     ("YAMATO",               "YAMATO"),
    #     ("JAPAN-SENKON",         "JAPAN-SENKON"),
    # )

    file = forms.FileField(
        label="Excelファイル（.xlsx）",
        help_text="* ヘッダー行を含むExcel（.xlsx）のみ対応",
        widget=forms.ClearableFileInput(attrs={
            "accept": ".xlsx",
        })
    )
    sheet_name = forms.CharField(
        label="シート名",
        required=False,
        help_text="未指定なら先頭シートを読み込みます。",
        widget=forms.TextInput(attrs={"placeholder": "例）ORDER"})
    )

    # destination = forms.ChoiceField(
    #     choices=(("", "-- 選択してください --"),) + DESTINATION_CHOICES,
    #     required=False,
    #     label="行先",
    # )
    # client = forms.ChoiceField(
    #     choices=(("", "-- 選択してください --"),) + CLIENT_CHOICES,
    #     required=False,
    #     label="取引先",
    # )