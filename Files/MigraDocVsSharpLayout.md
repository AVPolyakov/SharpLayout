## Сравнение производительности MigraDoc vs SharpLayout

Время создания pdf файла в зависимости от количества строк в таблице. MigraDoc – степенная функция с показателем 1.75. SharpLayout – линейная зависимость.  
![r8.png](MigraDocVsSharpLayout.png?raw=true)    
Измерения времени можно запустить, пройдя по ссылкам  
[для MigraDoc](https://github.com/AVPolyakov/MigraDoc-samples/blob/PerformanceLongTable/samples/core/HelloWorld/Program.cs#L16),  
[для SharpLayout](https://github.com/AVPolyakov/SharpLayout/blob/PerformanceLongTable/SharpLayout.Tests/Program.cs#L11).  
Таблица с результатами [ссылка](MigraDocVsSharpLayout.xlsx?raw=true).
