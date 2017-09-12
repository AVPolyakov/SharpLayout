## Live Viewer

Обновляем nuget пакет до версии SharpLayout.1.0.1

Компилируем проект [LiveViewer](LiveViewer/LiveViewer.csproj). В переменную окружения PATH добавляем директорию, в которой находится `LiveViewer.exe`.

С помощью метода [SavePng](SharpLayout.Tests/Program.cs#L17) сохраняем картинку на диск. Запускаем [StartLiveViewer](SharpLayout.Tests/Program.cs#L17). LiveViewer отслеживает изменения файла и автоматически обновляет изображение. Если метод `StartLiveViewer` вызывается с параметром `false`, то фокус остается в окне Visual Studio.

`Ctrl + Сlick` – навигация из отчета к строке исходного кода.

Для увеличения/уменьшения размера изображения следует изменить `resolution` в вызове метода `SavePng`.

[![Demo Video](Files/video.png?raw=true)](https://youtu.be/GOKvKWak8Kg)

Изменение размеров мышкой:

[![Изменение размеров мышкой](Files/video2.png?raw=true)](https://youtu.be/Zy6BkPnZxyY)

## Базовые элементы для создания отчета
**Таблица** представляет собой матрицу `N` на `M`:  
![Table.png](Files/Table.png?raw=true)  
В ячейку можно вставлять **параграф** или таблицу. Параграф состоит из коллекции **`Span` элементов**. Для `Span` элемента можно задать строку текста и параметры шрифта. Пример параграфа:  
![Paragraph.png](Files/Paragraph.png?raw=true)  

## Процесс создания отчета
[Скачиваем](http://www.consultant.ru/document/cons_doc_LAW_32449/6f63e20bf8ca001d7abacf60b7b29c8dfd44d261/)
из КонсультантПлюс форму платежного поручения (Форма 0401060) в формате MS-Word:  
![PaymentOrderBlank.png](Files/PaymentOrderBlank.png?raw=true)  
Создаем столбец `c1` и строку `r1`. Ширину столбца 3.5 см копируем из MS-Word:  
![Code1.png](Files/Code1.png?raw=true)  
На экране видим:  
![r1.png](Files/r1.png?raw=true)  
Добавляем в ячейку текст:  
![Code2.png](Files/Code2.png?raw=true)  
На экране отображается:  
![r2.png](Files/r2.png?raw=true)  
Выравниваем текст по центру:  
![Code3.png](Files/Code3.png?raw=true)  
![r3.png](Files/r3.png?raw=true)  
Добавляем нижний бордюр ячейки:  
![Code4.png](Files/Code4.png?raw=true)  
![r4.png](Files/r4.png?raw=true)  
Добавляем строку `r2` и вставляем текст в ячейку `r2[c1]`:  
![Code5.png](Files/Code5.png?raw=true)  
![r5.png](Files/r5.png?raw=true)  
Добавляем столбец `c2`. Ширину столбца 1.25 см копируем из MS-Word:  
![Code6.png](Files/Code6.png?raw=true")  
![r6.png](Files/r6.png?raw=true")  
Добавляем столбец `c3` и наполняем ячейки содержимым:  
![r7.png](Files/r7.png?raw=true")  
Ширину столбца `c5` 1.5 см копируем из MS-Word. Столбец `c4` должен занять все оставшееся на странице место. Поэтому ширину столбца `c4` зададим с помощью формулы:  
![Code7.png](Files/Code7.png?raw=true)  
Задание ширины с помощь формулы позволяет менять поля страницы и ширины других столбцов.  
![r8.png](Files/r8.png?raw=true")  
Аналогичным образом заполняем оставшиеся 63 ячейки.
В итоге получаем вот такое 
[платежное поручение](Files/PaymentOrder_Dev.pdf?raw=true).

## Платежное поручение
See the PDF file created by
[PaymentOrder.cs](SharpLayout.Tests/PaymentOrder.cs)
sample:
[PaymentOrder.pdf](Files/PaymentOrder.pdf?raw=true).
Highlighting of cells **r1c1** and paragraphs for development
[PaymentOrder_Dev.pdf](Files/PaymentOrder_Dev.pdf?raw=true)
![PaymentOrder.pdf](Files/PaymentOrder.png?raw=true")

## Справка о валютных операциях
[Ссылка](http://www.consultant.ru/document/cons_doc_LAW_133766/8408aeb59bc953ca3bbce8a729e5a5dca3bd0705/)
для скачивания формы справки о валютных операциях (ОКУД 0406009).
[Копия](Files/LAW191272_0_20170628_171359.RTF).  
C# код [Svo.cs](SharpLayout.Tests/Svo.cs) для создания справки о валютных операциях
[Svo.pdf](Files/Svo.pdf?raw=true). Подсветка в режиме разработки [Svo_Dev.pdf](Files/Svo_Dev.pdf?raw=true). Слева фиолетовым указаны номера строк в исходном коде [Svo.cs](SharpLayout.Tests/Svo.cs).  

## Номера C# строк для каждой ячейки
Для того чтобы увидеть номера C# строк для каждой ячейки необходимо установить флаг [IsHighlightCellLines](SharpLayout/Document.cs#L21). В файле [PaymentOrder_CellLines.pdf](Files/PaymentOrder_CellLines.pdf?raw=true)  в каждой ячейке в правом нижнем углу указан номер строки в файле [PaymentOrder.cs](SharpLayout.Tests/PaymentOrder.cs).