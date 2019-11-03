### Table of Contents  
- [Nuget package](#nuget-package)  
- [Base elements to create pdf document](#Base-elements-to-create-pdf-document)  
- [Examples](#Examples)  
  - [Payment Order](#Payment-Order)  
  - [Справка о валютных операциях](#Справка-о-валютных-операциях)  
  - [Паспорт сделки по контракту](#Паспорт-сделки-по-контракту)  
  - [Паспорт сделки по кредитному договору](#Паспорт-сделки-по-кредитному-договору)  
  - [Вставка векторной картинки](#Вставка-векторной-картинки)
- [Live Viewer](#live-viewer)  
- [Developer Tools](#developer-tools)
- [Data binding](#Data-binding)
- [Конвертация pdf файла в png для измерения расстояний](#Конвертация-pdf-файла-в-png-для-измерения-расстояний)  
- [Плагин к Paint.NET для измерения расстояний](#Плагин-к-paintnet-для-измерения-расстояний)
- [Сравнение производительности MigraDoc vs SharpLayout](#Сравнение-производительности-MigraDoc-vs-SharpLayout)
- [Процесс создания отчета](#Процесс-создания-отчета)  

## Nuget package
[Nuget package](https://www.nuget.org/packages/SharpLayout/)   

## Base elements to create pdf document
The table is a N × M matrix:  
![Table.png](Files/Table.png?raw=true)  
Paragraph, table and image can be inserted into cell of table.
Paragraph is a collection of `Span` elements.
Text and font parameters can be specified for `Span` element.
Paragraph exemple:  
![Paragraph.png](Files/Paragraph.png?raw=true)  

## Examples

### Payment Order
See the PDF file created by
[PaymentOrder.cs](Examples/PaymentOrder.cs)
sample:
[PaymentOrder.pdf](Files/PaymentOrder.pdf?raw=true).
Highlighting of cells **r1c1** and paragraphs for development
[PaymentOrder_Dev.pdf](Files/PaymentOrder_Dev.pdf?raw=true)
![PaymentOrder.pdf](Files/PaymentOrder.png?raw=true")

### Справка о валютных операциях
[Ссылка](http://www.consultant.ru/document/cons_doc_LAW_133766/8408aeb59bc953ca3bbce8a729e5a5dca3bd0705/)
для скачивания формы справки о валютных операциях (ОКУД 0406009).
[Копия](Files/LAW191272_0_20170628_171359.RTF).  
C# код [Svo.cs](Examples/Svo.cs) для создания справки о валютных операциях. Результат
[Svo.pdf](Files/Svo.pdf?raw=true). Подсветка в режиме разработки [Svo_Dev.pdf](Files/Svo_Dev.pdf?raw=true). Слева фиолетовым указаны номера строк в исходном коде [Svo.cs](Examples/Svo.cs).  

### Паспорт сделки по контракту
C# код [ContractDealPassport.cs](Examples/ContractDealPassport.cs) для создания паспорта сделки по контракту.
[ContractDealPassport.pdf](Files/ContractDealPassport.pdf?raw=true).  
[Ссылка](http://www.consultant.ru/document/Cons_doc_LAW_133766/775e60bb32004e2c078a5dca88b5bed4a0fa277e/)
для скачивания формы паспорта сделки по контракту (ОКУД 0406005 Форма 1).
[Копия](Files/LAW172722_0_20180026_141723.RTF).  

### Паспорт сделки по кредитному договору
C# код [LoanAgreementDealPassport.cs](Examples/LoanAgreementDealPassport.cs) для создания паспорта сделки по контракту.
[LoanAgreementDealPassport.pdf](Files/LoanAgreementDealPassport.pdf?raw=true).  
[Ссылка](http://www.consultant.ru/document/Cons_doc_LAW_133766/775e60bb32004e2c078a5dca88b5bed4a0fa277e/)
для скачивания формы паспорта сделки по контракту (ОКУД 0406005 Форма 2).
[Копия](Files/LAW172722_0_20180026_141723.RTF).  

### Вставка векторной картинки
См. метод `VectorImage` в файле [Tests.cs](Tests/Tests.cs). Результат [Rabbits.pdf](Files/Rabbits.pdf?raw=true).  

## Live Viewer

To use LiveViewer tool, you need to compile [LiveViewer](LiveViewer/LiveViewer.csproj) project. After compilation you need to add to PATH environment variable the directory path where `LiveViewer.exe` file is located.

SharpLayout can save a document as png image by method `SavePng`.

To open image file `Temp.png` by LiveViewer tool, you need run the following command:
```
LiveViewer Temp.png
```
LiveViewer tool tracks file changes and automatically updates a image on the screen. 

You can hold Ctrl and mouse left click on the image `Ctrl + Сlick`, LiveViewer will jump to the corresponding line of source code in Visual Studio or JetBrains Rider.

You can resize a elements by mouse. Select the size in Visual Studio editor. Hold left button and draw a rectangle by mouse in LiveViewer. Press `w` (width) or `h` (height) button. See second video below.

To change scale of image, change `resolution` parameter in `SavePng` method.

If you are running multiple instances of Visual Studio, you can specify a process PID. For example:  
```
LiveViewer Temp.png vs 15780
```

[![Demo Video](Files/video.png?raw=true)](https://youtu.be/GOKvKWak8Kg)

[![Изменение размеров мышкой](Files/video2.png?raw=true)](https://youtu.be/Zy6BkPnZxyY)

Navigation in JetBrains Rider:  
```
LiveViewer Temp.png rider
```
![SharpLayout_Rider.gif](Files/SharpLayout_Rider.gif?raw=true")

## Developer Tools
Отображение адресов ячеек R1C1 – `document.R1C1AreVisible = true`  
Подсветка ячеек – `document.CellsAreHighlighted = true`  
![r11.png](Files/r11.png?raw=true")  
Подсветка параграфов – `document.ParagraphsAreHighlighted = true`  
![r9.png](Files/r9.png?raw=true")  
Отображение номеров строк исходного кода в ячейке – `document.CellLineNumbersAreVisible = true`  
![r10.png](Files/r10.png?raw=true")  

## Data binding
Для того чтобы быстро проверить привязку данных можно вывести выражения в отчете `ExpressionVisible = true`:  
![r8.png](Files/r12.png?raw=true")  
Выражения отображаются в отчете, если привязка данных выполнена с помощью _expression tree lambda_:  
![r8.png](Files/r13.png?raw=true")  

## Конвертация pdf файла в png для измерения расстояний
 Программа [ConvertPDF](https://github.com/AVPolyakov/Pdf2Png) конвертирует любой pdf файл в png файл с разрешением 254 пикселя на дюйм. Таким образом, один пиксель соответствует одной десятой миллиметра.

 Компилируем проект [ConvertPDF](https://github.com/AVPolyakov/Pdf2Png). Запускаем ConvertPDF.exe. `Output format` выбираем `png256`.

 ## Плагин к Paint<span></span>.NET для измерения расстояний
 [Ссылка](https://github.com/AVPolyakov/MeasureSelection).

 ## Сравнение производительности MigraDoc vs SharpLayout
 [Ссылка](Files/MigraDocVsSharpLayout.md).

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