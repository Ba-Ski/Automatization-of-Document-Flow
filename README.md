# Automatization-of-Document-Flow

Инструменты:
 * C# - язык программирования
 * [EPPlus](https://github.com/JanKallman/EPPlus) - библиотека для работы с excel
 * [mongodb](https://www.mongodb.com/) - база данных.
 * [Serilog](https://serilog.net/) - логирование.
 * [Costura](https://github.com/Fody/Costura) - библиотека, помогающая собирать один исполняемый файл, в который встроены все зависимости.

Программы реализует следующий функционал:
 * Запись УП 2015 года в базу.
 * Расчёт доли ставки по дисциплине и остепенённости в ППС для УП 2015 года.
 * Записывает логи: об ошибках при чтении, об успешно прочитанных файлах (УП, ППС).

# Структура 

 * WindowsInterface - оконный пользовательский интерфейс, созданный с помощью Windiws Forms.
 * ConsoleInterface - консольный интерфейс.
 * СurriculumParse - библиотека классов. Содержит логику работы с excel и базой.
  * Директория Structures - содержит описание структур данных, которые записываются в базу данных.
    * Curriculum - структура, описывающая учебный план шаблона 2015 года. Нужно внести изменения, чтобы хранить УП шаблонов всех годов.
    * Subject - структура, описывающая предмет из учебного плана.
    * Complexity - структура, описывающая трудоемкость определённого предмета. 
  
# Использование

Программа умеет читать учебные планы 2015 года и складывать их в базу данных, также рассчитвывать долю ставки и остепенённость для соответствующих ППС.

Чтобы рассчитать долю ставки и остепенённость определённого ППС (excel-файл), необходимо вначале записать в базу данных соответствующий учебный план.
Для это можно по кнопке "Загрузить все УП2015" загрузить все доступные учебные программы, расположенные в указаной вами директории.
В этом случае должна соблюдаться следующая иерархия директорий:

Верхняя директория (База ОПОП)
* директории - номера специальностей
  * директории - года программ
    * файлы соответствующие программе (УП, ППС итд.).

Программа попытается прочитать все учебные программы, а о своих успехах напишет в файлы "Ошибки" и "Прочитанные файлы".
В файле "Ошибки" будут указаны файлы, которые не удалось прочитать, и директории, в которых не удалось найти файлов.
В файле "Прочитанные файлы" будет список успешно прочитанных файлов.

Также можно загрузить в базу один файл, нажав кнопку "Загрузить один УП2015".

После того, как необходимые УП были загружены, можно рассчитывать соответствующие им ППС-файлы по кнопке "Рассчитать ППС2015". После расчёта программа сообщит о результатах.
В указанный ППС-файл будет добавлен один столбец с рассчитанными долями ставки для каждой строки таблицы. В последнюю ячейку данного столбца будет записана остепенённсоть в виде формулы excel.
Часто не удаётся рассчитать долю ставки для всех строк в ППС-файле. Это связано с тем, что ППС или УП были заполнено некорректно или неединообразно. Для ППС-файла это, например, отсутствие вида занятия для определённого предмета.
Часто бывает, что индексы предметов в УП и ППС не совпадают (например, лишняя точка после последней цифры индекса).
