from django import forms
from datetime import datetime

class UploadFileForm(forms.Form):
    file = forms.FileField(
        label='Выберите ZIP-архив', 
        required=True,
        widget=forms.ClearableFileInput(attrs={'class': 'input'}) 
    )
    existing_excel = forms.FileField(
        label='Существующий Excel (необязательно)', 
        required=False,
        widget=forms.ClearableFileInput(attrs={'class': 'input'}) 
    )
    
    date = forms.DateField(
        label='Дата',
        required=False,
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'input'}) 
    )
    tn_ved_code = forms.CharField(
        label='Код ТН ВЭД', 
        initial='27132000',
        required=True,
        widget=forms.TextInput(attrs={'class': 'input'}) 
    )
    bnd_code = forms.CharField(
        label='БНД', 
        initial='60/90',
        required=True,
        widget=forms.TextInput(attrs={'class': 'input'})
    )
    nds_percent = forms.DecimalField(
        label='НДС ЕАЭС (%)', 
        initial=12,
        decimal_places=2,
        min_value=0,
        required=True,
        widget=forms.NumberInput(attrs={'class': 'input'}) 
    )
    save_photos = forms.BooleanField(
        label='Сохранить фотографии',
        initial=False,
        required=False
    )

    def clean_file(self):
        file = self.cleaned_data.get('file')
        if file:
            if not file.name.lower().endswith('.zip'):
                raise forms.ValidationError('Пожалуйста, загрузите файл с расширением .zip')
            
            if file.content_type != 'application/zip' and file.content_type != 'application/x-zip-compressed':
                 pass
        return file

    def clean_existing_excel(self):
        file = self.cleaned_data.get('existing_excel')
        if file:
            valid_extensions = ['.xlsx', '.xlsm', '.xltx', '.xltm']
            if not any(file.name.lower().endswith(ext) for ext in valid_extensions):
                raise forms.ValidationError('Пожалуйста, загрузите валидный Excel файл (.xlsx, .xlsm, .xltx, .xltm)')
        return file


class PreviewEditForm(forms.Form):
    """Динамическая форма для редактирования данных предпросмотра"""
    
    def __init__(self, *args, **kwargs):
        objects_data = kwargs.pop('objects_data', [])
        super().__init__(*args, **kwargs)
        
        for idx, obj_data in enumerate(objects_data):
            prefix = f'obj_{idx}'

            date_initial = obj_data.get(1, '')
            date_iso = ''
            if date_initial:
                try:
                    if isinstance(date_initial, str):
                        parsed_date = datetime.strptime(date_initial, '%d.%m.%Y').date()
                    else:
                        parsed_date = date_initial
                    date_iso = parsed_date.strftime('%Y-%m-%d')
                except Exception:
                    date_iso = ''

            self.fields[f'{prefix}_date'] = forms.DateField(
                label='Дата',
                required=False,
                input_formats=['%Y-%m-%d', '%d.%m.%Y'],
                initial=date_iso,
                widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'})
            )
            
            self.fields[f'{prefix}_marka'] = forms.CharField(
                label='Марка АТС',
                initial=obj_data.get(2, ''),
                required=False,
                widget=forms.TextInput(attrs={'class': 'form-control'})
            )
            
            self.fields[f'{prefix}_gos_number'] = forms.CharField(
                label='Гос.номер',
                initial=obj_data.get(3, ''),
                required=False,
                widget=forms.TextInput(attrs={'class': 'form-control'})
            )
            
            self.fields[f'{prefix}_fio'] = forms.CharField(
                label='ФИО водителя',
                initial=obj_data.get(4, ''),
                required=False,
                widget=forms.TextInput(attrs={'class': 'form-control'})
            )

            self.fields[f'{prefix}_kol_ton'] = forms.CharField(
                label='Кол.тон',
                initial=obj_data.get(7, ''),
                required=False,
                widget=forms.TextInput(attrs={'class': 'form-control'})
            )
            
            price_value = obj_data.get(8)
            if price_value:
                price_str = str(price_value) if not isinstance(price_value, str) else price_value
            else:
                price_str = ''
            self.fields[f'{prefix}_price'] = forms.DecimalField(
                label='Цена',
                initial=price_str,
                required=False,
                decimal_places=2,
                widget=forms.NumberInput(attrs={'class': 'form-control', 'step': '0.01'})
            )
            
            date_sopr = obj_data.get(13, '')
            self.fields[f'{prefix}_date_sopr'] = forms.CharField(
                label='Дата сопр.накл',
                initial=date_sopr,
                required=False,
                widget=forms.TextInput(attrs={'class': 'form-control'})
            )
            
            num_sopr = obj_data.get(15, '')
            self.fields[f'{prefix}_num_sopr'] = forms.CharField(
                label='№ сопров.накл. KZ',
                initial=num_sopr,
                required=False,
                widget=forms.TextInput(attrs={'class': 'form-control'})
            )
            
            invoice = obj_data.get(16, '')
            self.fields[f'{prefix}_invoice'] = forms.CharField(
                label='№ счет факт (Инвойс)',
                initial=invoice,
                required=False,
                widget=forms.TextInput(attrs={'class': 'form-control'})
            )
