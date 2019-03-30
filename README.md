### Table of Contents  
- [Nuget package](#nuget-package)  
- [Base elements to create pdf document](#Base-elements-to-create-pdf-document)  
- [Процесс создания отчета](#%D0%9F%D1%80%D0%BE%D1%86%D0%B5%D1%81%D1%81-%D1%81%D0%BE%D0%B7%D0%B4%D0%B0%D0%BD%D0%B8%D1%8F-%D0%BE%D1%82%D1%87%D0%B5%D1%82%D0%B0)  
- [Примеры](#%D0%9F%D1%80%D0%B8%D0%BC%D0%B5%D1%80%D1%8B)  
  - [Платежное поручение](#%D0%9F%D0%BB%D0%B0%D1%82%D0%B5%D0%B6%D0%BD%D0%BE%D0%B5-%D0%BF%D0%BE%D1%80%D1%83%D1%87%D0%B5%D0%BD%D0%B8%D0%B5)  
  - [Справка о валютных операциях](#%D0%A1%D0%BF%D1%80%D0%B0%D0%B2%D0%BA%D0%B0-%D0%BE-%D0%B2%D0%B0%D0%BB%D1%8E%D1%82%D0%BD%D1%8B%D1%85-%D0%BE%D0%BF%D0%B5%D1%80%D0%B0%D1%86%D0%B8%D1%8F%D1%85)  
  - [Паспорт сделки по контракту](#%D0%9F%D0%B0%D1%81%D0%BF%D0%BE%D1%80%D1%82-%D1%81%D0%B4%D0%B5%D0%BB%D0%BA%D0%B8-%D0%BF%D0%BE-%D0%BA%D0%BE%D0%BD%D1%82%D1%80%D0%B0%D0%BA%D1%82%D1%83)  
  - [Паспорт сделки по кредитному договору](#%D0%9F%D0%B0%D1%81%D0%BF%D0%BE%D1%80%D1%82-%D1%81%D0%B4%D0%B5%D0%BB%D0%BA%D0%B8-%D0%BF%D0%BE-%D0%BA%D1%80%D0%B5%D0%B4%D0%B8%D1%82%D0%BD%D0%BE%D0%BC%D1%83-%D0%B4%D0%BE%D0%B3%D0%BE%D0%B2%D0%BE%D1%80%D1%83)  
  - [Вставка векторной картинки](#%D0%B2%D1%81%D1%82%D0%B0%D0%B2%D0%BA%D0%B0-%D0%B2%D0%B5%D0%BA%D1%82%D0%BE%D1%80%D0%BD%D0%BE%D0%B9-%D0%BA%D0%B0%D1%80%D1%82%D0%B8%D0%BD%D0%BA%D0%B8)
- [Live Viewer](#live-viewer)  
- [Developer Tools](#developer-tools)
- [Привязка данных](#%D0%9F%D1%80%D0%B8%D0%B2%D1%8F%D0%B7%D0%BA%D0%B0-%D0%B4%D0%B0%D0%BD%D0%BD%D1%8B%D1%85)
- [Конвертация pdf файла в png для измерения расстояний](#%D0%9A%D0%BE%D0%BD%D0%B2%D0%B5%D1%80%D1%82%D0%B0%D1%86%D0%B8%D1%8F-pdf-%D1%84%D0%B0%D0%B9%D0%BB%D0%B0-%D0%B2-png-%D0%B4%D0%BB%D1%8F-%D0%B8%D0%B7%D0%BC%D0%B5%D1%80%D0%B5%D0%BD%D0%B8%D1%8F-%D1%80%D0%B0%D1%81%D1%81%D1%82%D0%BE%D1%8F%D0%BD%D0%B8%D0%B9)  
- [Плагин к Paint.NET для измерения расстояний](#%D0%9F%D0%BB%D0%B0%D0%B3%D0%B8%D0%BD-%D0%BA-paintnet-%D0%B4%D0%BB%D1%8F-%D0%B8%D0%B7%D0%BC%D0%B5%D1%80%D0%B5%D0%BD%D0%B8%D1%8F-%D1%80%D0%B0%D1%81%D1%81%D1%82%D0%BE%D1%8F%D0%BD%D0%B8%D0%B9)
- [Сравнение производительности MigraDoc vs SharpLayout](#%D0%A1%D1%80%D0%B0%D0%B2%D0%BD%D0%B5%D0%BD%D0%B8%D0%B5-%D0%BF%D1%80%D0%BE%D0%B8%D0%B7%D0%B2%D0%BE%D0%B4%D0%B8%D1%82%D0%B5%D0%BB%D1%8C%D0%BD%D0%BE%D1%81%D1%82%D0%B8-migradoc-vs-sharplayout)

## Nuget package
[Nuget package](https://www.nuget.org/packages/SharpLayout/)   


## Base elements to create pdf document
The table is a N × M matrix:  
![Table.png](Files/Table.png?raw=true)  
Paragraph, table and picture can be inserted into cell of table.
Paragraph is a collection of `Span` elements.
Text and font parameters can be specified for `Span` element.
Paragraph exemple:  
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

## Примеры

### Платежное поручение
See the PDF file created by
[PaymentOrder.cs](SharpLayout.Tests/PaymentOrder.cs)
sample:
[PaymentOrder.pdf](Files/PaymentOrder.pdf?raw=true).
Highlighting of cells **r1c1** and paragraphs for development
[PaymentOrder_Dev.pdf](Files/PaymentOrder_Dev.pdf?raw=true)
![PaymentOrder.pdf](Files/PaymentOrder.png?raw=true")

### Справка о валютных операциях
[Ссылка](http://www.consultant.ru/document/cons_doc_LAW_133766/8408aeb59bc953ca3bbce8a729e5a5dca3bd0705/)
для скачивания формы справки о валютных операциях (ОКУД 0406009).
[Копия](Files/LAW191272_0_20170628_171359.RTF).  
C# код [Svo.cs](SharpLayout.Tests/Svo.cs) для создания справки о валютных операциях. Результат
[Svo.pdf](Files/Svo.pdf?raw=true). Подсветка в режиме разработки [Svo_Dev.pdf](Files/Svo_Dev.pdf?raw=true). Слева фиолетовым указаны номера строк в исходном коде [Svo.cs](SharpLayout.Tests/Svo.cs).  

### Паспорт сделки по контракту
C# код [ContractDealPassport.cs](SharpLayout.Tests/ContractDealPassport.cs) для создания паспорта сделки по контракту.
[ContractDealPassport.pdf](Files/ContractDealPassport.pdf?raw=true).  
[Ссылка](http://www.consultant.ru/document/Cons_doc_LAW_133766/775e60bb32004e2c078a5dca88b5bed4a0fa277e/)
для скачивания формы паспорта сделки по контракту (ОКУД 0406005 Форма 1).
[Копия](Files/LAW172722_0_20180026_141723.RTF).  

### Паспорт сделки по кредитному договору
C# код [LoanAgreementDealPassport.cs](SharpLayout.Tests/LoanAgreementDealPassport.cs) для создания паспорта сделки по контракту.
[LoanAgreementDealPassport.pdf](Files/LoanAgreementDealPassport.pdf?raw=true).  
[Ссылка](http://www.consultant.ru/document/Cons_doc_LAW_133766/775e60bb32004e2c078a5dca88b5bed4a0fa277e/)
для скачивания формы паспорта сделки по контракту (ОКУД 0406005 Форма 2).
[Копия](Files/LAW172722_0_20180026_141723.RTF).  

### Вставка векторной картинки
См. метод `VectorImage` в файле [Tests.cs](SharpLayout.Tests/Tests.cs). Результат [Rabbits.pdf](Files/Rabbits.pdf?raw=true).  

## Live Viewer

Компилируем проект [LiveViewer](LiveViewer/LiveViewer.csproj). В переменную окружения PATH добавляем директорию, в которой находится `LiveViewer.exe`.

С помощью метода [SavePng](SharpLayout.Tests/Program.cs#L17) сохраняем картинку на диск. Запускаем [StartLiveViewer](SharpLayout.Tests/Program.cs#L17). LiveViewer отслеживает изменения файла и автоматически обновляет изображение. Если метод `StartLiveViewer` вызывается с параметром `false`, то фокус остается в окне Visual Studio.

`Ctrl + Сlick` – навигация из отчета к строке исходного кода.

Изменение размеров мышкой  – см. второе видео.

Для увеличения/уменьшения размера изображения следует изменить `resolution` в вызове метода `SavePng`.

Если запущено несколько экземпляров Visual Studio, то можно указать PID процесса. Например,  
`LiveViewer.exe Temp.png vs 15780`  

[![Demo Video](Files/video.png?raw=true)](https://youtu.be/GOKvKWak8Kg)

[![Изменение размеров мышкой](Files/video2.png?raw=true)](https://youtu.be/Zy6BkPnZxyY)

Навигация в JetBrains Rider:  
![SharpLayout_Rider.gif](Files/SharpLayout_Rider.gif?raw=true")

## Developer Tools
Отображение адресов ячеек R1C1 – `document.R1C1AreVisible = true`  
Подсветка ячеек – `document.CellsAreHighlighted = true`  
![r11.png](Files/r11.png?raw=true")  
Подсветка параграфов – `document.ParagraphsAreHighlighted = true`  
![r9.png](Files/r9.png?raw=true")  
Отображение номеров строк исходного кода в ячейке – `document.CellLineNumbersAreVisible = true`  
![r10.png](Files/r10.png?raw=true")  

## Привязка данных
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