using System;
using System.Data;
using System.Diagnostics;
using System.Net;
using System.Text;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using Validata.PKI;

namespace WSAlamedaApiSample
{
    internal class Program
    {
        // имя криптопрофиля
        private const string CryptoProfile = "TEST";
        // депозитарный код
        private const string ParticipantCode = "XXXXXXXXXXXX";
        // адрес веб-сервиса
        private const string BaseUrl = "https://gost-g.nsd.ru/WSAlamedags/";
        private const string DetlUrl = "/v2/participants";
        private const string Method = "/deals";

        private static void Main(string[] args)
        {
            Console.WriteLine("Загрузка данных...");

            // Фомрирование сроки для подписи
            var stringToSign = String.Format("{0}/{1}{2}", DetlUrl, ParticipantCode, Method);
            var filter = "";
//            filter = String.Format("?serv_code={0}", "RMBC");
            var requestUrl = String.Format("{0}{1}{2}", BaseUrl, stringToSign, filter);

            var wc = new WebClient {Encoding = Encoding.UTF8};
            // Подписывание строки для авторизации при обращении к веб-сервису
            var sign = SignString(stringToSign);
            wc.Headers.Add("X-WSALAM-SIGN", sign);
            var json = wc.DownloadString(requestUrl);

            var deals = JObject.Parse(json);

            Console.WriteLine("Формирование Excel...");

            var dealsTable = PrepareDealsDataTable("Deals");

            foreach (var deal in deals["dealPool"]["deals"])
            {
                AddDealRow(dealsTable, deal);
            }
            var wb = new XLWorkbook();

            wb.Worksheets.Add(dealsTable);

            var fileName = String.Format("{0}-{1:dd.MM.yyyy}.xlsx", deals["doc_no"], deals["doc_date_time"]);

            wb.SaveAs(fileName);
            Console.WriteLine("Готово.");

            Process.Start(fileName);
            Console.ReadKey();
        }

        private static DataTable PrepareDealsDataTable(string tableName)
        {
            var table = new DataTable {TableName = tableName};
            table.Columns.Add("DEAL_NUM", typeof (string));
            table.Columns.Add("DEAL_DATE", typeof(DateTime));
            table.Columns.Add("OBLIGATION_TYPE", typeof(string));
            table.Columns.Add("FIRST_PART_DT", typeof(DateTime));
            table.Columns.Add("SECOND_PART_DT", typeof(DateTime));
            table.Columns.Add("PART2_VOLUME", typeof(string));
            table.Columns.Add("REQUEST_VOLUME", typeof(string));
            table.Columns.Add("LIABILITY_CURR", typeof(string));
            table.Columns.Add("REPO_RATE", typeof(string));
            table.Columns.Add("DEBT_CODE", typeof(string));
            table.Columns.Add("CRED_CODE", typeof(string));
            table.Columns.Add("OBLIG_SUM", typeof(string));
            table.Columns.Add("LIAB_SUM", typeof(string));
            table.Columns.Add("LIAB", typeof(string));

            table.Columns.Add("COLLATERAL_TYPE", typeof(string));
            table.Columns.Add("THRESHOLD", typeof(string));
            table.Columns.Add("PLACE_OF_TRADE", typeof(string));
            table.Columns.Add("INT_METH", typeof(string));
            table.Columns.Add("SERV_CODE", typeof(string));
            table.Columns.Add("RESULT", typeof(string));
            return table;
        }

        private static void AddDealRow(DataTable dealsTable, JToken deal)
        {
            dealsTable.Rows.Add(
                deal["deal_num"],
                deal["deal_date"],
                deal["obligation_type"],
                deal["first_part_dt"],
                deal["second_part_dt"],
                deal["part2_volume"],
                deal["request_volume"],
                deal["liability_curr"],
                deal["repo_rate"],
                deal["debt_code"],
                deal["cred_code"],
                deal["oblig_sum"],
                deal["liab_sum"],
                deal["liab"],

                deal["collateral_type"],
                deal["threshold"],
                deal["place_of_trade"],
                deal["int_meth"],
                deal["serv_code"],
                deal["result"]
                );
        }

        private static String SignString(String stringToSign)
        {
            var data = Encoding.UTF8.GetBytes(stringToSign);
            VcertObject ctx = new VcertObject();
            try
            {
                // инициализируем криптографию (имененм профиля)
                ctx.Initialize(CryptoProfile, VcertObject.InitializeFlags.NoCrlUpdate | VcertObject.InitializeFlags.NoLdap | VcertObject.InitializeFlags.UseRegistry);
                // определяем параметры подписи
                SignParameters sp = new SignParameters
                {
                    Detached = true,
                    Pkcs7 = true
                };
                StreamSignCtx signContext = new StreamSignCtx(sp, null);
                // подписываем массив байт
                byte[] result = ctx.StreamSignMem(signContext, data, true);
                // конвертируем подпись в Base64
                return Convert.ToBase64String(result);
            }
            catch (VcertException e)
            {
                Console.WriteLine("Error - {0}", e.Message);
                throw new ApplicationException("Error Initializing main context", e);
            }
        }

    }
}