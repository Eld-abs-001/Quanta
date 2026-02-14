import os
import shutil
from decimal import Decimal, ROUND_HALF_UP
from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.conf import settings
from .forms import UploadFileForm, PreviewEditForm
from .services import get_current_dollar_rate, process_zip_file, generate_excel, NetworkError

@login_required
def upload_view(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            try:
                date = form.cleaned_data.get('date')
                tn_ved_code = form.cleaned_data['tn_ved_code']
                bnd_code = form.cleaned_data['bnd_code']
                nds_percent = form.cleaned_data['nds_percent']
                save_photos = form.cleaned_data['save_photos']
                existing_excel = request.FILES.get('existing_excel')
                
                dollar_rate = Decimal('0')
                
                if date:
                    date_str = date.strftime('%d.%m.%Y')
                    print(f"[upload_view] Received upload. date_str={date_str}, tn_ved_code={tn_ved_code}, bnd_code={bnd_code}, nds_percent={nds_percent}, save_photos={save_photos}, existing_excel={bool(existing_excel)}")
                    
                    try:
                        dollar_rate = get_current_dollar_rate(date_str)
                        print(f"[upload_view] Got dollar_rate={dollar_rate} for date {date_str}")
                    except NetworkError as e:
                        print(f"[upload_view] NetworkError getting dollar rate for {date_str}: {e.technical_details}")
                        full_message = f"{e.user_message}|||{e.technical_details}"
                        messages.error(request, full_message)
                        return render(request, 'work/index.html', {'form': form})
                    except Exception as e:
                        error_message = str(e)
                        print(f"[upload_view] Error getting dollar rate for {date_str}: {error_message}")
                        messages.error(request, f'Ошибка при получении курса доллара: {error_message}')
                        return render(request, 'work/index.html', {'form': form})
                else:
                    print(f"[upload_view] Received upload without date. Skipping dollar rate fetch.")

                
                request.session['saved_defaults'] = {
                    'date': date.isoformat() if date else None,
                    'tn_ved_code': tn_ved_code,
                    'bnd_code': bnd_code,
                    'nds_percent': float(nds_percent),
                    'save_photos': save_photos
                }
                
                try:
                    print(f"[upload_view] Calling process_zip_file with file={getattr(request.FILES['file'], 'name', None)}")
                    results = process_zip_file(
                        request.FILES['file'],
                        dollar_rate=dollar_rate,
                        selected_date=date,
                        tn_ved_code=tn_ved_code,
                        bnd_code=bnd_code,
                        nds_percent=nds_percent,
                        save_photos=save_photos
                    )
                    print(f"[upload_view] process_zip_file returned {len(results)} result(s)")
                except Exception as e:
                    error_message = str(e)
                    print(f"[upload_view] Error processing zip file: {error_message}")
                    messages.error(request, f'Ошибка при обработке файла: {error_message}')
                    return render(request, 'work/index.html', {'form': form})
                
                has_critical_errors = False
                for row in results:
                    if row.get('errors'):
                        has_critical_errors = True
                        driver_name = row.get(4, 'Неизвестный водитель')
                        for error in row['errors']:
                            messages.error(request, f"{driver_name}: {error}")
                
                if has_critical_errors:
                    return render(request, 'work/index.html', {'form': form})
                
                existing_excel_path = None
                if existing_excel:
                    base_temp_dir = os.path.join(settings.MEDIA_ROOT, "temp_ocr")
                    existing_excel_dir = os.path.join(base_temp_dir, "existing_excel")
                    os.makedirs(existing_excel_dir, exist_ok=True)
                    existing_excel_path = os.path.join(existing_excel_dir, existing_excel.name)
                    with open(existing_excel_path, 'wb+') as destination:
                        for chunk in existing_excel.chunks():
                            destination.write(chunk)
                
                serializable_results = []
                for row in results:
                    serializable_row = {}
                    for key, value in row.items():
                        if key == 'preview_images':
                            serializable_row[key] = value
                        elif key == 'field_images':
                            field_images_serialized = {}
                            if isinstance(value, dict):
                                for field_key, img_list in value.items():
                                    field_images_serialized[str(field_key)] = img_list
                            serializable_row[key] = field_images_serialized
                        elif isinstance(value, Decimal):
                            serializable_row[str(key)] = str(value)
                        else:
                            serializable_row[str(key)] = value
                    serializable_results.append(serializable_row)
                
                request.session['preview_data'] = {
                    'results': serializable_results,
                    'dollar_rate': str(dollar_rate),
                    'tn_ved_code': tn_ved_code,
                    'bnd_code': bnd_code,
                    'nds_percent': str(nds_percent),
                    'existing_excel_path': existing_excel_path,
                    'save_photos': save_photos
                }
                
                return redirect('preview')
                
            except Exception as e:
                error_message = str(e)
                messages.error(request, f'Произошла ошибка: {error_message}')
                return render(request, 'work/index.html', {'form': form})
    else:
        initial_data = request.session.get('saved_defaults', {})
        if 'date' in initial_data:
            try:
                from datetime import datetime as dt
                initial_data['date'] = dt.fromisoformat(initial_data['date']).date()
            except (ValueError, TypeError):
                initial_data.pop('date', None)
        form = UploadFileForm(initial=initial_data)
    
    return render(request, 'work/index.html', {'form': form})


@login_required
def preview_view(request):
    preview_data = request.session.get('preview_data')
    
    if not preview_data:
        messages.error(request, 'Данные для предпросмотра не найдены. Пожалуйста, загрузите файл заново.')
        return redirect('upload')
    
    results = preview_data['results']
    
    form = PreviewEditForm(objects_data=results)
    
    objects_for_template = []
    for idx, row in enumerate(results):
        image_paths = row.get('preview_images', [])
        valid_images = []
        for img_path in image_paths:
            full_path = os.path.join(settings.MEDIA_ROOT, img_path)
            if os.path.exists(full_path):
                valid_images.append(img_path)
        
        data_dict = {}
        field_images_dict = {}
        sources_dict = {}
        
        for key_str, value in row.items():
            if key_str == 'preview_images':
                continue
            elif key_str == 'field_images':
                if isinstance(value, dict):
                    for field_key_str, img_list in value.items():
                        try:
                            field_key = int(field_key_str)
                            valid_field_images = []
                            for img_path in img_list:
                                full_path = os.path.join(settings.MEDIA_ROOT, img_path)
                                if os.path.exists(full_path):
                                    valid_field_images.append(img_path)
                            if valid_field_images:
                                field_images_dict[field_key] = valid_field_images
                        except (ValueError, TypeError):
                            pass
                continue
            elif key_str == 'sources':
                if isinstance(value, dict):
                    for source_key_str, source_val in value.items():
                        if str(source_key_str) == '1':
                            continue
                        try:
                            sources_dict[int(source_key_str)] = source_val
                        except (ValueError, TypeError):
                            sources_dict[source_key_str] = source_val
                continue
            try:
                key = int(key_str)
                data_dict[key] = value
            except (ValueError, TypeError):
                data_dict[key_str] = value
        
        obj = {
            'index': idx,
            'images': valid_images,
            'data': data_dict,
            'field_images': field_images_dict,
            'sources': sources_dict,
            'errors': row.get('errors', [])
        }
        date_iso = ""
        date_raw = data_dict.get(1)
        if date_raw:
            try:
                if hasattr(date_raw, 'strftime'):
                    date_iso = date_raw.strftime('%Y-%m-%d')
                else:
                    from datetime import datetime as dt
                    parsed = dt.strptime(str(date_raw), '%d.%m.%Y').date()
                    date_iso = parsed.strftime('%Y-%m-%d')
            except Exception:
                date_iso = ""
        obj['date_iso'] = date_iso
        objects_for_template.append(obj)
    
    context = {
        'form': form,
        'objects': objects_for_template,
        'media_url': settings.MEDIA_URL
    }
    
    return render(request, 'work/preview.html', context)

@login_required
def preview_submit_view(request):
    if request.method != 'POST':
        return redirect('preview')
    
    action = request.POST.get('action', 'ready')
    print(f"[preview_submit_view] Called. action={action}")
    
    preview_data = request.session.get('preview_data')
    
    if not preview_data:
        messages.error(request, 'Данные для предпросмотра не найдены. Пожалуйста, загрузите файл заново.')
        return redirect('upload')
    
    results = preview_data['results']
    
    form = PreviewEditForm(request.POST, objects_data=results)
    
    if form.is_valid():
        print(f"[preview_submit_view] Form is valid. preview_data keys: {list(preview_data.keys())}, results_count={len(results)}")
        def serialize_results(rows):
            serialized = []
            for row in rows:
                row_ser = {}
                for key, value in row.items():
                    key_str = str(key)
                    if key_str == 'field_images' and isinstance(value, dict):
                        field_images_serialized = {str(k): v for k, v in value.items()}
                        row_ser[key_str] = field_images_serialized
                        continue
                    if key_str == 'sources' and isinstance(value, dict):
                        sources_serialized = {str(k): v for k, v in value.items()}
                        row_ser[key_str] = sources_serialized
                        continue
                    if key_str == 'preview_images':
                        row_ser[key_str] = value
                        continue
                    if key_str == 'errors':
                        row_ser[key_str] = value
                        continue
                    if isinstance(value, Decimal):
                        row_ser[key_str] = str(value)
                    else:
                        row_ser[key_str] = value
                serialized.append(row_ser)
            return serialized

        updated_results = []
        has_rate_errors = False

        for idx, row in enumerate(results):
            prefix = f'obj_{idx}'
            
            updated_row = {}
            
            for key_str, value in row.items():
                if key_str == 'preview_images':
                    updated_row['preview_images'] = value
                    continue
                if key_str == 'field_images':
                    updated_row['field_images'] = value
                    continue
                if key_str == 'sources':
                    updated_row['sources'] = value
                    continue
                if key_str == 'errors':
                    updated_row['errors'] = value
                    continue
                
                try:
                    key = int(key_str)
                except (ValueError, TypeError):
                    continue
                
                if key in [7, 8, 9, 10, 11, 12] and isinstance(value, str):
                    try:
                        updated_row[key] = Decimal(value)
                    except:
                        updated_row[key] = value
                else:
                    updated_row[key] = value
            
            date_value = form.cleaned_data.get(f'{prefix}_date')
            if not date_value:
                raw_date = request.POST.get(f'{prefix}_date')
                if raw_date:
                    try:
                        from datetime import datetime as dt
                        parsed_date = dt.strptime(raw_date, '%Y-%m-%d').date()
                        date_value = parsed_date
                    except Exception:
                        try:
                            parsed_date = dt.strptime(raw_date, '%d.%m.%Y').date()
                            date_value = parsed_date
                        except Exception:
                            date_value = None
            if date_value:
                updated_row[1] = date_value.strftime('%d.%m.%Y')
            else:
                updated_row[1] = updated_row.get(1, '')

            marka_value = form.cleaned_data.get(f'{prefix}_marka', '')
            updated_row[2] = str(marka_value).strip() if marka_value else (updated_row.get(2, '') or '')
            
            gos_number_value = form.cleaned_data.get(f'{prefix}_gos_number', '')
            updated_row[3] = str(gos_number_value).strip() if gos_number_value else (updated_row.get(3, '') or '')
            
            fio_value = form.cleaned_data.get(f'{prefix}_fio', '')
            updated_row[4] = str(fio_value).strip() if fio_value else (updated_row.get(4, '') or '')
            
            kol_ton_value = form.cleaned_data.get(f'{prefix}_kol_ton', '')
            if kol_ton_value:
                try:
                    updated_row[7] = Decimal(str(kol_ton_value).replace(',', '.'))
                except:
                    updated_row[7] = updated_row.get(7, Decimal("0"))
            
            price_value = form.cleaned_data.get(f'{prefix}_price')
            if price_value is not None:
                updated_row[8] = Decimal(str(price_value))
            
            date_sopr_value = form.cleaned_data.get(f'{prefix}_date_sopr', '')
            updated_row[13] = str(date_sopr_value).strip() if date_sopr_value else (updated_row.get(13, '') or '')
            
            num_sopr_value = form.cleaned_data.get(f'{prefix}_num_sopr', '')
            updated_row[15] = str(num_sopr_value).strip() if num_sopr_value else (updated_row.get(15, '') or '')
            
            invoice_value = form.cleaned_data.get(f'{prefix}_invoice', '')
            updated_row[16] = str(invoice_value).strip() if invoice_value else (updated_row.get(16, '') or '')
            
            for key in [1, 2, 3, 4, 5, 6, 13, 14, 15, 16]:
                if key not in updated_row:
                    updated_row[key] = ''
            
            try:
                kol_ton = updated_row.get(7, Decimal("0"))
                if isinstance(kol_ton, str):
                    kol_ton = Decimal(kol_ton)
                cena = updated_row.get(8, Decimal("0"))
                if isinstance(cena, str):
                    cena = Decimal(cena)

                base_rate = updated_row.get(10, None)
                if base_rate is None:
                    base_rate = preview_data.get('dollar_rate', '0')
                try:
                    base_rate = Decimal(base_rate)
                except Exception:
                    base_rate = Decimal("0")

                rate_to_use = base_rate
                
                original_date_str = row.get('1')
                current_date_str = updated_row.get(1)
                
                date_changed = False
                if original_date_str != current_date_str:
                    date_changed = True
                    print(f"[preview_submit_view] Date changed for obj {idx}: {original_date_str} -> {current_date_str}")

                if 'errors' in updated_row and updated_row['errors']:
                    updated_row['errors'] = [e for e in updated_row['errors'] if not str(e).startswith("Ошибка при получении курса")]

                if action == 'recalculate' or date_changed:
                    try:
                        date_for_rate = updated_row.get(1)
                        if not date_for_rate or str(date_for_rate).lower() == 'none':
                            pass 
                        else:
                            print(f"[preview_submit_view] Recalculate/DateChanged: fetching rate for date {date_for_rate}")
                            rate_to_use = get_current_dollar_rate(date_for_rate)
                            print(f"[preview_submit_view] Got rate {rate_to_use} for date {date_for_rate}")
                    except NetworkError as e:
                        rate_error = f"{e.user_message}|||{e.technical_details}"
                        print(f"[preview_submit_view] NetworkError for date {date_for_rate}: {e.technical_details}")
                        updated_row.setdefault('errors', [])
                        updated_row['errors'].append(rate_error)
                        has_rate_errors = True
                        rate_to_use = base_rate
                    except Exception as e:
                        rate_error = f"Ошибка при получении курса на дату {updated_row.get(1, '')}: {e}"
                        print(f"[preview_submit_view] get_current_dollar_rate exception for date {date_for_rate}: {e}")
                        updated_row.setdefault('errors', [])
                        updated_row['errors'].append(rate_error)
                        has_rate_errors = True
                        rate_to_use = base_rate

                updated_row[10] = rate_to_use

                nds_percent = Decimal(preview_data['nds_percent'])
                
                sum_dollar = (kol_ton * cena).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                updated_row[9] = sum_dollar
                
                sum_som = (sum_dollar * rate_to_use).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                updated_row[11] = sum_som
                
                nds_sum = (sum_som * nds_percent / Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                updated_row[12] = nds_sum
            except Exception as e:
                print(f"[preview_submit_view] Calculation error in preview for obj {idx}: {e}")
            
            updated_results.append(updated_row)

        if action == 'recalculate':
            if has_rate_errors:
                messages.error(request, 'Не удалось получить курс доллара для некоторых строк при перерасчёте. Исправьте дату или введите курс вручную.')
            else:
                messages.success(request, 'Перерасчёт выполнен.')

            preview_data['results'] = serialize_results(updated_results)
            request.session['preview_data'] = preview_data
            print(f"[preview_submit_view] Recalculate finished. Saved {len(updated_results)} rows to session. Redirecting to preview.")
            return redirect('preview')
        
        existing_excel = None
        existing_excel_path = preview_data.get('existing_excel_path')
        if existing_excel_path and os.path.exists(existing_excel_path):
            existing_excel = existing_excel_path
        
        try:
            print(f"[preview_submit_view] Generating Excel for {len(updated_results)} rows. existing_excel={bool(existing_excel_path)}")
            excel_data = []
            for row_idx, row in enumerate(updated_results):
                excel_row = {}
                for key in range(1, 19):
                    if key in row:
                        value = row[key]
                        if key in [7, 8, 9, 10, 11, 12] and isinstance(value, str):
                            try:
                                value = Decimal(value)
                            except:
                                pass
                        if value is None:
                            if key in [1, 2, 3, 4, 5, 6, 13, 14, 15, 16, 17, 18]:
                                value = ''
                        excel_row[key] = value
                    else:
                        if key in [1, 2, 3, 4, 5, 6, 13, 14, 15, 16, 17, 18]:
                            excel_row[key] = ''
                        else:
                            excel_row[key] = None
                
                excel_data.append(excel_row)
            
            nds_percent = preview_data.get('nds_percent', 2)
            
            wb = generate_excel(excel_data, existing_excel, nds_percent=nds_percent)
            
            save_photos = preview_data.get('save_photos', False)
            base_temp_dir = os.path.join(settings.MEDIA_ROOT, "temp_ocr")
            preview_imgs_dir = os.path.join(base_temp_dir, "preview_imgs")
            
            if save_photos:
                if os.path.exists(preview_imgs_dir):
                    shutil.rmtree(preview_imgs_dir)
                if os.path.exists(base_temp_dir):
                    for item in ['upload', 'extracted', 'existing_excel']:
                        item_path = os.path.join(base_temp_dir, item)
                        if os.path.exists(item_path):
                            shutil.rmtree(item_path)
            else:
                if os.path.exists(base_temp_dir):
                    shutil.rmtree(base_temp_dir)
            
            if 'preview_data' in request.session:
                del request.session['preview_data']
            
            response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename=ocr_results.xlsx'
            wb.save(response)
            
            messages.success(request, 'Excel файл успешно создан и загружен!')
            
            return response
            
        except Exception as e:
            error_message = str(e)
            print(f"[preview_submit_view] Exception while generating Excel: {error_message}")
            messages.error(request, f'Ошибка при создании Excel файла: {error_message}')
            return redirect('preview')
    else:
        messages.error(request, 'Пожалуйста, исправьте ошибки в форме.')
        return redirect('preview')


def custom_page_not_found_view(request, exception):
    return render(request, "errors/error_404.html", status=404)

def custom_permission_denied_view(request, exception=None):
    return render(request, "errors/error_403.html", status=403)