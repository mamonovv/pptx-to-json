### Конвертер уммет работать с Текстом, Картинками и таблицами

## Как он работает

### Во-первых, код создает папку под картинки images в том месте, где был запущен файл.
### Во-вторых, если говорить о структуре Json, то она такая:

    {
    'Slide{номер по счету}': {
        'Text{номер по счету}': ...,
        'Image{номер по счету}': ...,
        'Table{номер по счету}': {
            'row{номер по счету}': [...],
            'row{номер по счету}': [...],
         },
    'Slide{номер по счету}': ...
    }

### То есть идут слайды, а потом внутри идет контент. Номер по счету идет относительно родителя. Слайды считаются в целом на весь файл, Text,Image,Table для слайда, а row для таблицы.

