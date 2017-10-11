# Инструкция по подключению к WSAlameda API с помощью приложения на C#

## Введение

В данном примере демонстрируется создание консольного приложения с помощью MS Visual Studio.

Приложение:
* подключается к API WSAlameda
* запрашивает данные об открытых сделках
* формирует Excel файл с полученными данными.

![screen1](Screenshots/excel_doc.png?raw=true "Информация по открытым сделкам")

## Подготовка проекта

Каждый из запросов к веб-сервиу должен быть подписан. Подписывается часть строки запроса, результат кодируется в формате base64 и помещается в http заголовок сообщения "X-WSALAM-SIGN".
Работа с криптографией осуществляется через API предоставляемое АПК Клиент МБ (ПК "Справочник сертификатов"). Для работы с ним необходимо запросить описание API и обертку для C# в технической поддержке ЭДО Московской Биржи (http://www.moex.com/s1314)
После получения библиотеки (vcpia2.dll), ссылка на нее должна быть добавлена в созданный проект.

Создаем новое консольное приложение в Visual Studio
File->New->Project Выбираем Console Application

Взаимодействие с API WSAlameda (возвращаемый результат) осуществляется в формате JSON.
С форматом JSON можно работать как с обычным текстом, одно из его основных преимуществ - "человекочитаемость".

Пример возвращаемых данных API:

```json
{
  "doc_no": "1480406",
  "doc_date_time": "2017-10-09T18:24:49.792+03:00",
  "dealPool": {
    "oblig_sum_p": "0",
    "liab_sum_p": "0",
    "reval_level": "0",
    "liab_p": "FLAT",
    "deals":[
      {
        "deal_num": "TLMBK09101",
        "reg_num": "30824",
        "deal_date": "2017-10-09T00:00:00+03:00",
        "main_order_reg_num": "1106131",
        "counter_order_reg_num": "1106132",
        "dtl_status": "DEAL",
        "status": "1",
        "obligation_type": "NONE",
        "place_of_trade": "BLM",
        "first_part_dt": "2017-10-09T00:00:00+03:00",
        "second_part_dt": "2017-10-20T00:00:00+03:00",
        "part2_volume": "5099.45",
        "request_volume": "5000",
        "real_volume": "5000",
        "repo_rate": "66",
        "threshold": "10",
        "reuse": "Y",
        "without_suo": "0",
        "shift_term_dt": "1",
        "return_var": "1",
        "securities": [
          {
            "security_code": "SU46018RMFS6",
            "name": "TEST",
            "quantity": "5",
            "discount": "0",
            "price_types_priority": "BL"
          }
        ],
        "deb_bnk_acc": "30411810100002000757",
        "deb_bic": "MICURUMMXXX",
        "deb_dep_acc": "TS170914001B",
        "deb_dep_sect": "00000000000000000",
        "deb_dep_sect_id": "11338676",
        "cred_bnk_acc": "30411810500002000043",
        "cred_bic": "MICURUMMXXX",
        "cred_dep_acc": "TS1710020066",
        "cred_dep_sect": "00000000000000000",
        "cred_dep_sect_id": "11394218",
        "oblig_sum": "5000",
        "liab_sum": "6000",
        "liab": "EXCS",
        "result": "DEAL",
        "collateral_type": "B0000000075J",
        "collateral_id": "5343691460",
        "int_meth": "365/366",
        "deal_calc_type": "DVP1",
        "deal_calc_type2": "DVP3",
        "debt_code": "MC0138600000",
        "cred_code": "MC0010100000",
        "related": "N",
        "serv_code": "RMBC",
        "liability_curr": "RUB",
        "f_calc_ext": "0",
        "f_calc_ext_complete": "0"
      }
    ]
  }
}
```

Однако рекомендуем использовать для работы с форматом JSON специализированную библиотеку. 
Например [Newtonsoft Json.NET](http://www.newtonsoft.com/json)
Библиотека позволяет выполнять сериализацию (преобразование объектов .NET в формат JSON) и десериализовывать JSON документы в привычные структуры .NET (типизированные объекты, перечисления, массивы и т.д. )
Это позволит сильно сократить количество программного кода и позволит в полной мере использовать средства платформы .net для работы с данными, в том числе LINQ

Для подключения пакеты из NuGet репозитория, открываем [Package Manager Console](https://docs.nuget.org/docs/start-here/using-the-package-manager-console) и выполняем команду

```
PM> Install-Package Newtonsoft.Json
```

Так как в данном примере данные будут выгружаться в Excel, подключим пакет [ClosedXML](https://github.com/closedxml/closedxml)
```
PM> Install-Package ClosedXML
```

С его помощью двумя строчками кода можно создать Excel документ и помесить на лист данные.



В секцию using файла Program.cs добавляем пространства имен, которые нам понадобятся:

```cs
	using System;
	using System.Data;
	using System.Diagnostics;
	using System.Net;
	using System.Text;
	using ClosedXML.Excel;
	using Newtonsoft.Json.Linq;
	using Validata.PKI;
```

В них мы будем использовать следующие классы:

Пространство имен  | Класс | Сценарий использования
--- | --- | ---
System.Data  | DataTable  | Таблица, для сохранения результатов
System.Diagnostics  | Process  | Открыть файл в Excel
System.Linq  |   | Методы расширения для работы с коллекциями
System.Net  | WebClient  | Выполнения HTTP запросов
System.Text  | Encoding  | Указать кодировку документа
ClosedXML.Excel  | XLWorkbook  | Формирования файла в формате Excel
Newtonsoft.Json  |   | Сериализация/Десериализация JSON документов
Newtonsoft.Json.Linq  |   | LINQ to JSON
Validata.PKI  | VcertObject, SignParameters, StreamSignCtx | Подписание строки запроса для авторизации в веб-сервисе

## Подключение к API WSAlameda

Для получения данных в API WSAlameda используется метод GET протокола HTTP. 

Формат запроса для GET методов:

http://t4-wl.test.local:7510/WSAlamedaT4/v2/participants/{participant}/<метод>?<фильтр>&<фильтр>...


Объявляем константы, которые будем испльовать в коде

```cs
	// имя криптопрофиля
	private const string CryptoProfile = "TEST";
	// депозитарный код
	private const string ParticipantCode = "MC0010100000";
	// адрес веб-сервиса
	private const string BaseUrl = "http://t4-wl.test.local:7510/WSAlamedaT4";
	private const string DetlUrl = "/v2/participants";
	private const string Method = "/deals";
```

Формируем строку подключения к API

```cs
	// Фомрирование сроки для подписи
	var stringToSign = String.Format("{0}/{1}{2}", DetlUrl, ParticipantCode, Method);
	var filter = ""; // String.Format("?serv_code={0}", "RMBC");
	var requestUrl = String.Format("{0}{1}{2}", BaseUrl, stringToSign, filter);
```

Создаем WebClient, подписываем сроку запроса, результат добавляем в http header и выполняем запрос к веб-сервису

```cs
	var wc = new WebClient {Encoding = Encoding.UTF8};
	// Подписывание строки для авторизации при обращении к веб-сервису
	var sign = SignString(stringToSign);
	wc.Headers.Add("X-WSALAM-SIGN", sign);
	var json = wc.DownloadString(requestUrl);
```

В ответ получена строка, содержащая JSON документ. Десериализуем ее в массив JObject

```cs
	var deals = JObject.Parse(json);
```

Подготавливаем объект DataTable для экспорта его в Excel:

```cs
	var dealsTable = PrepareDealsDataTable("Deals");
```

Итерируемся по списку сделок и заполняем таблицы

```cs
	foreach (var deal in deals["dealPool"]["deals"])
	{
	    AddDealRow(dealsTable, deal);
	}
```

Создаем новый документ Excel, добавляем в него лист с информацией о сделках, сохраняем на диск и открываем в приложении


```cs
	var wb = new XLWorkbook();
	
	wb.Worksheets.Add(dealsTable);
	
	var fileName = String.Format("{0}-{1:dd.MM.yyyy}.xlsx", deals["doc_no"], deals["doc_date_time"]);
	
	wb.SaveAs(fileName);
	Console.WriteLine("Готово.");
	
	Process.Start(fileName);
	Console.ReadKey();
```
