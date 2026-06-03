FIELDS_MAP_TYPE_1 = {
    "Дата (1)": (880, 2485, 390, 140),
    "ФИО Водит. (4)": (850, 2395, 500, 150),
    "Кол.тон (7)": (1470, 1355, 300, 80),
    "Марка": (350, 2655, 200, 70),
    "Гос_номер ()": (-80, 2645, 470, 170),
    "Якорь (1)": (0, 500, 500, 1000),
}

FIELDS_MAP_TYPE_2 = {
    "Цена (8)": (1585, 2160, 250, 160),
    "№ счет факт (Инвойс) (16)": (485, 190, 500, 60),
}

FIELDS_MAP_TYPE_3 = {
    "№ сопров.накл. KZ (15)": (360, 250, 330, 200),
    "Дата сопр.накл (13)": (360, 250, 330, 200)
}

FIELDS_MAP_TYPE_2_PAGE_2 = {
    "Цена (8) Alt": (1500, 100, 150, 120),
}

MIN_HEIGHT_CONFIG = {
    "ФИО Водит. (4)": 20,
    "Марка_Гос_номер ()": 38,
}

CURRENT_MIN_HEIGHT_CONFIG = MIN_HEIGHT_CONFIG

LEGACY_FIELDS_MAP_TYPE_1 = {
    "Дата (1)": (880, 2520, 390, 140),
    "ФИО Водит. (4)": (850, 2450, 500, 150),
    "Кол.тон (7)": (1500, 1390, 300, 80),
    "Марка": (380, 2730, 250, 70),
    "Гос_номер ()": (-80, 2730, 470, 170),
    "Якорь (1)": (0, 500, 500, 1000),
}

LEGACY_FIELDS_MAP_TYPE_2 = {
    "Цена (8)": (1450, 2150, 250, 160),
    "№ счет факт (Инвойс) (16)": (500, 200, 500, 60),
}

LEGACY_FIELDS_MAP_TYPE_3 = {
    "№ сопров.накл. KZ (15)": (360, 250, 330, 200),
    "Дата сопр.накл (13)": (360, 250, 330, 200)
}

LEGACY_FIELDS_MAP_TYPE_2_PAGE_2 = {
    "Цена (8) Alt": (1500, 100, 150, 120),
}

LEGACY_MIN_HEIGHT_CONFIG = {
    "ФИО Водит. (4)": 20,
    "Марка_Гос_номер ()": 38,
}






{% load static %}
<!DOCTYPE html>
<html lang="ru">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Предпросмотр данных</title>
    <link rel="icon" href="{% static 'work/icons/icon.png' %}">
    <link rel="stylesheet" href="{% static 'work/css/style.css' %}">

    <style>
        .btn-update-download,
        .btn-cancel {
            display: none;
        }

        .btn-update-download.visible,
        .btn-cancel.visible {
            display: inline-block;
        }

        .btn-ready.hidden {
            display: none;
        }

        .field-input-group label {
            font-weight: bold;
        }
        .debug-section {
            width: 42%;
            min-width: 360px;
            border-left: 1px solid #e5e7eb;
            padding-left: 1rem;
        }

        .debug-section h4 {
            margin: 0 0 0.75rem 0;
        }

        .debug-field-block {
            margin-bottom: 0.9rem;
            padding: 0.6rem;
            border-radius: 8px;
            background: #f8fafc;
            border: 1px solid #e2e8f0;
        }

        .debug-field-title {
            font-weight: 600;
            margin-bottom: 0.5rem;
        }

        .debug-item {
            font-size: 0.9rem;
            line-height: 1.35;
            margin-bottom: 0.2rem;
        }

        .debug-item.ok {
            color: #065f46;
        }

        .debug-item.filtered {
            color: #b91c1c;
        }

        @media (max-width: 768px) {
            .debug-section {
                width: 100%;
                min-width: 0;
                border-left: none;
                border-top: 1px solid #e5e7eb;
                padding-left: 0;
                padding-top: 1rem;
            }
        }
    </style>

    <script>
        document.addEventListener('DOMContentLoaded', function () {
            const form = document.querySelector('form');
            const updateBtn = document.querySelector('.btn-update-download');
            const cancelBtn = document.querySelector('.btn-cancel');
            const readyBtn = document.querySelector('.btn-ready');
            let formChanged = false;

            const formInputs = form.querySelectorAll('input[type="text"], input[type="number"], input[type="date"]');

            const initialValues = new Map();
            formInputs.forEach(input => {
                initialValues.set(input.name, input.value);
            });

            function checkFormChanged() {
                let hasChanges = false;
                formInputs.forEach(input => {
                    if (input.value !== initialValues.get(input.name)) {
                        hasChanges = true;
                    }
                });

                if (hasChanges && !formChanged) {
                    formChanged = true;
                    updateBtn.classList.add('visible');
                    cancelBtn.classList.add('visible');
                    readyBtn.classList.add('hidden');
                } else if (!hasChanges && formChanged) {
                    formChanged = false;
                    updateBtn.classList.remove('visible');
                    cancelBtn.classList.remove('visible');
                    readyBtn.classList.remove('hidden');
                }
            }

            function resetForm() {
                formInputs.forEach(input => {
                    const initialValue = initialValues.get(input.name);
                    input.value = initialValue;
                });
                checkFormChanged();
            }

            formInputs.forEach(input => {
                input.addEventListener('input', checkFormChanged);
                input.addEventListener('change', checkFormChanged);
            });

            if (cancelBtn) {
                cancelBtn.addEventListener('click', function (event) {
                    event.preventDefault();
                    resetForm();
                });
            }

            if (form) {
                form.addEventListener('submit', function (event) {
                    if (event.submitter && event.submitter.value === 'ready') {
                        setTimeout(function () {
                            window.location.href = '{% url "upload" %}';
                        }, 1000);
                    }
                });
            }
        });
    </script>
</head>

<body class="preview-page">
    <div class="container preview-container">
        <h1>Предпросмотр данных</h1>

        {% if messages %}
        {% for message in messages %}
        <div class="{{ message.tags }}">
            <strong>
                {% if message.tags == 'success' %}
                Успешно:
                {% elif message.tags == 'error' %}
                Ошибка:
                {% else %}
                Сообщение:
                {% endif %}
            </strong> {{ message }}
        </div>
        {% endfor %}
        {% endif %}
        <div class="drivers-nav">
            <div class="drivers-nav-title">Имя водителей</div>
            <div class="drivers-list">
                {% for obj in objects %}
                    <a href="#driver_{{ forloop.counter0 }}" class="driver-link {% if obj.data.plate_format_warning %}driver-link-warning{% endif %}">
                        {{ obj.data.4|default:"Не указано" }}
                    </a>
                {% endfor %}
            </div>
        </div>

        <form method="post" action="{% url 'preview_submit' %}">
            {% csrf_token %}

            {% for obj in objects %}
            <div class="object-container" id="driver_{{ forloop.counter0 }}">
                <div class="form-section">
                    <div class="object-title">({{ forloop.counter }}) {{ obj.data.4 }}</div>

                    {% if obj.errors %}
                    <div class="error">
                        <strong>Обнаружены проблемы:</strong>
                        <ul style="margin: 0.5rem 0 0 1.5rem; padding: 0;">
                            {% for error in obj.errors %}
                            <li>{{ error }}</li>
                            {% endfor %}
                        </ul>
                    </div>
                    {% endif %}

                    <div class="form-section-group">
                        <div class="form-section-group-title">Данные с отсканированного документа</div>

                        <div class="field-with-image">
                            <div class="field-input-group">
                                <label>Дата:</label>
                                <input type="date" name="obj_{{ forloop.counter0 }}_date"
                                    value="{{ obj.date_iso|default:'' }}" class="form-control"
                                    data-initial-date="{{ obj.date_iso|default:'' }}">
                            </div>
                        </div>

                        <div class="field-with-image">
                            {% if obj.field_images.2 %}
                            <div class="field-image-group">
                                {% for img_path in obj.field_images.2 %}
                                <div class="image-container">
                                    <img src="{{ media_url }}{{ img_path }}" alt="Марка АТС">
                                </div>
                                {% endfor %}
                            </div>
                            {% elif obj.sources.2 %}
                            <div class="field-image-group">
                                <div class="excel-source">
                                    Данные из ячейки {{ obj.sources.2 }}
                                </div>
                            </div>
                            {% endif %}
                            <div class="field-input-group">
                                <label>Марка АТС:</label>
                                <input type="text" name="obj_{{ forloop.counter0 }}_marka"
                                    value="{{ obj.data.2|default:'' }}" class="form-control">
                            </div>
                        </div>

                        <div class="field-with-image">
                            {% if obj.field_images.3 %}
                            <div class="field-image-group">
                                {% for img_path in obj.field_images.3 %}
                                <div class="image-container">
                                    <img src="{{ media_url }}{{ img_path }}" alt="Гос.номер">
                                </div>
                                {% endfor %}
                            </div>
                            {% elif obj.sources.3 %}
                            <div class="field-image-group">
                                <div class="excel-source">
                                    Данные из ячейки {{ obj.sources.3 }}
                                </div>
                            </div>
                            {% endif %}
                            <div class="field-input-group">
                                <label>Гос.номер:</label>
                                <input type="text" name="obj_{{ forloop.counter0 }}_gos_number"
                                    value="{{ obj.data.3|default:'' }}" class="form-control">
                            </div>
                        </div>

                        <div class="field-with-image">
                            {% if obj.field_images.4 %}
                            <div class="field-image-group">
                                {% for img_path in obj.field_images.4 %}
                                <div class="image-container">
                                    <img src="{{ media_url }}{{ img_path }}" alt="ФИО водителя">
                                </div>
                                {% endfor %}
                            </div>
                            {% elif obj.sources.4 %}
                            <div class="field-image-group">
                                <div class="excel-source">
                                    Данные из ячейки {{ obj.sources.4 }}
                                </div>
                            </div>
                            {% endif %}
                            <div class="field-input-group">
                                <label>ФИО водителя:</label>
                                <input type="text" name="obj_{{ forloop.counter0 }}_fio"
                                    value="{{ obj.data.4|default:'' }}" class="form-control">
                            </div>
                        </div>

                        <div class="field-with-image">
                            {% if obj.field_images.7 %}
                            <div class="field-image-group">
                                {% for img_path in obj.field_images.7 %}
                                <div class="image-container">
                                    <img src="{{ media_url }}{{ img_path }}" alt="Кол.тон">
                                </div>
                                {% endfor %}
                            </div>
                            {% elif obj.sources.7 %}
                            <div class="field-image-group">
                                <div class="excel-source">
                                    Данные из ячейки {{ obj.sources.7 }}
                                </div>
                            </div>
                            {% endif %}
                            <div class="field-input-group">
                                <label>Кол.тон:</label>
                                <input type="text" name="obj_{{ forloop.counter0 }}_kol_ton"
                                    value="{{ obj.data.7|default:'' }}" class="form-control">
                            </div>
                        </div>

                        <div class="field-with-image">
                            {% if obj.field_images.8 %}
                            <div class="field-image-group">
                                {% for img_path in obj.field_images.8 %}
                                <div class="image-container">
                                    <img src="{{ media_url }}{{ img_path }}" alt="Цена">
                                </div>
                                {% endfor %}
                            </div>
                            {% elif obj.sources.8 %}
                            <div class="field-image-group">
                                <div class="excel-source">
                                    Данные из ячейки {{ obj.sources.8 }}
                                </div>
                            </div>
                            {% endif %}
                            <div class="field-input-group">
                                <label>Цена:</label>
                                <input type="number" name="obj_{{ forloop.counter0 }}_price"
                                    value="{{ obj.data.8|default:'' }}" step="0.01" class="form-control">
                            </div>
                        </div>

                        <div class="field-with-image">
                            {% if obj.field_images.13 %}
                            <div class="field-image-group">
                                {% for img_path in obj.field_images.13 %}
                                <div class="image-container">
                                    <img src="{{ media_url }}{{ img_path }}" alt="Дата сопр.накл">
                                </div>
                                {% endfor %}
                            </div>
                            {% elif obj.sources.13 %}
                            <div class="field-image-group">
                                <div class="excel-source">
                                    Данные из ячейки {{ obj.sources.13 }}
                                </div>
                            </div>
                            {% endif %}
                            <div class="field-input-group">
                                <label>Дата сопр.накл:</label>
                                <input type="text" name="obj_{{ forloop.counter0 }}_date_sopr"
                                    value="{{ obj.data.13|default:'' }}" class="form-control">
                            </div>
                        </div>

                        <div class="field-with-image">
                            {% if obj.field_images.15 %}
                            <div class="field-image-group">
                                {% for img_path in obj.field_images.15 %}
                                <div class="image-container">
                                    <img src="{{ media_url }}{{ img_path }}" alt="№ сопров.накл. KZ">
                                </div>
                                {% endfor %}
                            </div>
                            {% endif %}
                            <div class="field-input-group">
                                <label>№ сопров.накл. KZ:</label>
                                <input type="text" name="obj_{{ forloop.counter0 }}_num_sopr"
                                    value="{{ obj.data.15|default:'' }}" class="form-control">
                            </div>
                        </div>

                        <div class="field-with-image">
                            {% if obj.field_images.16 %}
                            <div class="field-image-group">
                                {% for img_path in obj.field_images.16 %}
                                <div class="image-container">
                                    <img src="{{ media_url }}{{ img_path }}" alt="№ счет факт">
                                </div>
                                {% endfor %}
                            </div>
                            {% endif %}
                            <div class="field-input-group">
                                <label>№ счет факт (Инвойс):</label>
                                <input type="text" name="obj_{{ forloop.counter0 }}_invoice"
                                    value="{{ obj.data.16|default:'' }}" class="form-control">
                            </div>
                        </div>

                    </div>
                </div>
                <div class="debug-section">
                    <h4>Все OCR данные (с размерами)</h4>
                    {% if obj.data.ocr_debug %}
                        {% for field_name, items in obj.data.ocr_debug.items %}
                        <div class="debug-field-block">
                            <div class="debug-field-title">{{ field_name }}</div>
                            {% for item in items %}
                            <div class="debug-item {% if item.passed %}ok{% else %}filtered{% endif %}">
                                {{ item.text }} | H: {{ item.height }} | Min: {{ item.min_height }} |
                                {% if item.passed %}OK{% else %}FILTERED{% endif %}
                            </div>
                            {% endfor %}
                        </div>
                        {% endfor %}
                    {% else %}
                    <div class="debug-field-block">
                        Нет OCR debug данных для этого документа.
                    </div>
                    {% endif %}
                </div>
            </div>
            {% endfor %}


            <div class="submit-button-container" style="display: flex; justify-content: center;">
                <div style="display: flex; gap: 1rem; justify-content: center; width: 75%;">
                    <button type="button" class="btn-submit btn-cancel" style="background-color: #6c757d;">Отменить изменения</button>
                    <button type="submit" name="action" value="recalculate" class="btn-submit btn-update-download"
                    style="background-color: #f0ad4e;">Обновить</button>
                    <button type="submit" name="action" value="ready" class="btn-submit btn-ready">Готов</button>
                </div>
            </div>
        </form>
    </div>
</body>

</html>




if any(x in field_name for x in ["ФИО Водит.", "Марка", "Гос_номер"]):
                r, g, b = crop_img.split()
                crop_img = b 
                
                if "ФИО Водит." in field_name:
                    threshold = fio_threshold
                    crop_img = crop_img.point(lambda p: 255 if p > threshold else 0)
                elif "Гос_номер" in field_name:
                    threshold = 165
                    threshold_options = [165, 155, 175, 145, 185, 135, 195]
                    best_results = []
                    successful_results = []
                    selected_threshold = threshold_options[0]
                    min_height_plate = CURRENT_MIN_HEIGHT_CONFIG.get(field_name, 0)

                    if reader is not None:
                        for idx, threshold_try in enumerate(threshold_options):
                            try:
                                candidate_img = crop_img.point(lambda p: 255 if p > threshold_try else 0)
                                candidate_np = np.array(candidate_img)
                                candidate_results = reader.readtext(candidate_np, detail=1)
                                if idx == 0:
                                    best_results = candidate_results
                                    selected_threshold = threshold_try

                                candidate_raw_items = []
                                for (bbox, text, prob) in candidate_results:
                                    height = int(((bbox[3][1] - bbox[0][1]) + (bbox[2][1] - bbox[1][1])) / 2)
                                    if height >= min_height_plate:
                                        candidate_raw_items.append((text, height))

                                if DataCleaner.is_plate_result_like(candidate_raw_items):
                                    successful_results = candidate_results
                                    selected_threshold = threshold_try
                                    break
                            except Exception as retry_err:
                                print(f"[extract_text_from_pdf] Plate retry OCR error threshold={threshold_try}: {retry_err}")

                    if successful_results:
                        results = successful_results
                    else:
                        results = best_results
                        extracted_data["_plate_retry_failed"] = True
                        print(
                            f"[extract_text_from_pdf] Plate format did not match expected pattern after 7 attempts. "
                            f"Using first attempt threshold={selected_threshold}."
                        )
                    crop_img = crop_img.point(lambda p: 255 if p > selected_threshold else 0)
                    plate_ocr_precomputed = True
            
            img_filename = get_safe_filename(pdf_path, field_name)
            img_path = os.path.join(save_dir, img_filename)
            crop_img.save(img_path)