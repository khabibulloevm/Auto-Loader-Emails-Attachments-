using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using System.Text;
using EmailReseiver.Contexts;
using EmailReseiver.Models;
using EmailReseiver.DAL;
using MailKit;
using Microsoft.Extensions.Configuration;
using MimeKit;
using MessageSummaryItems = MailKit.MessageSummaryItems;
using Bytescout.Spreadsheet;
using System.Text.RegularExpressions;
using EmailReseiver.Services;
using Newtonsoft.Json;
using System.Xml.Serialization;
using Microsoft.VisualBasic;

namespace EmailReseiver.MailServices
{
    public class MailReceiverService
    {
        public RecipeEItemList recipes { get; set; }
        public IConfiguration Configuration { get; }
        public List<MailItem> listOfMessages = new List<MailItem>();
        public MailReceiverService()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(System.AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json",
                    optional: true,
                    reloadOnChange: true);
            Configuration = builder.Build();

            //Строка подключения к БД (находится в appsettings.json)
            var connectionString = Configuration.GetConnectionString("LocalSql");

            IServiceCollection services = new ServiceCollection();
            services.AddDbContext<Context>(options => options.UseSqlServer(connectionString));
            services.AddScoped<ImportDataService>();
            services.AddScoped<ImportDataDuplicateService>();
            var provider = services.BuildServiceProvider().CreateScope();
            _importDataService = provider.ServiceProvider.GetRequiredService<ImportDataService>();
            _doublesService = provider.ServiceProvider.GetRequiredService<ImportDataDuplicateService>();
        }

        //Заполнение WorkSupplierDogovorId, зависимо от FinancingItem
        public string getLetter(string financingItem)
        {
            if (Regex.IsMatch(financingItem, "Смеш", RegexOptions.IgnoreCase)) return "201";
            if (Regex.IsMatch(financingItem, "регион", RegexOptions.IgnoreCase)) return "202";
            if (Regex.IsMatch(financingItem, "федера", RegexOptions.IgnoreCase)) return "203";

            return "";
        }

        //Преобразование СНИЛС в единой стандартной форме 
        private string castSnilsValue(string SNILS)
        {
            if (String.IsNullOrEmpty(SNILS)) return SNILS;

            try
            {
                string snils_numbers_only = SNILS.Trim().Replace(" ", "").Replace("-", "");
                if (Information.IsNumeric(snils_numbers_only) && snils_numbers_only.Length == 11)
                {
                    return $"{snils_numbers_only.Substring(0, 3)}-{snils_numbers_only.Substring(3, 3)}-{snils_numbers_only.Substring(6, 3)} {snils_numbers_only.Substring(9, 2)}";
                }
                else
                {
                    return string.Empty;
                }
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
        }


        //Информация о лигине и пароль почтового ящика находится в appsettings.json
        public async Task<List<MailItem>> DoReceiveMail()
        {
            var yandexUser = Configuration["YandexUser"];
            var yandexPass = Configuration["YandexPass"];

            try
            {
                while (true)
                {

                    using (var client = new MailKit.Net.Imap.ImapClient())
                    {
                        await client.ConnectAsync("imap.yandex.ru", 993, true);
                        await client.AuthenticateAsync(yandexUser, yandexPass);

                        await client.Inbox.OpenAsync(MailKit.FolderAccess.ReadWrite);

                        var uids = await client.Inbox.SearchAsync(MailKit.Search.SearchQuery.NotSeen);

                        var messages = await client.Inbox.FetchAsync(uids,
                            MessageSummaryItems.Envelope | MessageSummaryItems.BodyStructure);

                        if (messages != null && messages.Count > 0)
                        {
                            foreach (var msg in messages)
                            {
                                client.Inbox.AddFlags(uids, MailKit.MessageFlags.Seen, true);

                                listOfMessages.Add(new MailItem
                                {
                                    Date = msg.Date.ToString(),
                                    From = msg.Envelope.From.ToString(),
                                    Subj = msg.Envelope.Subject,
                                    HasAttachments = msg.Attachments != null && msg.Attachments.Count() > 0,
                                });

                                foreach (var att in msg.Attachments.OfType<BodyPartBasic>())
                                {
                                    var part = (MimePart)await client.Inbox.GetBodyPartAsync(msg.UniqueId, att);

                                    if (Regex.IsMatch(part.FileName, "XLSX") || Regex.IsMatch(part.FileName, "XLS")) continue;


                                    Stream outStream = new MemoryStream();
                                    await part.Content.DecodeToAsync(outStream);
                                    outStream.Position = 0;
                                    Spreadsheet document = new Spreadsheet();
                                    document.LoadFromStream(outStream);
                                    var sheet = document.Workbook.Worksheets[0];

                                    //Проверка пустых строк в начале документа 
                                    int rowIndex = 0;

                                    for (int row = 0; row <= sheet.Rows.LastFormatedRow; row++)
                                    {
                                        if (sheet.Cell(row, 0).ValueAsString != "")
                                        {
                                            rowIndex = row + 1;
                                            break;
                                        }
                                    }

                                    //Проверка пустых строк в середине документа 
                                    for (int row = rowIndex; row <= sheet.Rows.LastFormatedRow; row++)
                                    {

                                        if (sheet.Cell(row, 0).ValueAsString == "")
                                        {
                                            for (int row_1 = 0; row_1 < 1000; row_1++)
                                            {
                                                if (sheet.Cell(row_1, 0).ValueAsString != "")
                                                {
                                                    row += row_1;
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            try
                                            {
                                                //Процесс извлечения инфрмации из API и запись в БД 
                                                ImportData importData = getData(sheet, row);

                                               

                                                

                                                ImportData impD_2 = importData;

                                                recipes = new EmailReseiver.DAL.RecipeEItemList();


                                                if (convertedData.IsError)
                                                {
                                                    throw new Exception(convertedData.ErrorText);
                                                }
                                                else
                                                {

                                                    var serializer = new XmlSerializer(typeof(EmailReseiver.DAL.RecipeEItemList));
                                                    using (TextReader reader = new StringReader((string)convertedData.Data))
                                                    {
                                                        recipes = (RecipeEItemList)serializer.Deserialize(reader);
                                                    }
                                                }

                                                
                                                foreach (var recipe in recipes.it)
                                                {
                                                    if (recipe.Number == importData.RecNum)
                                                    {
                                                        impD_2.LastName = recipe.LastName;
                                                        impD_2.Name = recipe.FirstName;
                                                        impD_2.MidName = recipe.FatherName;
                                                        impD_2.RecSeria = recipe.Serial;
                                                        impD_2.MKB = recipe.MKBCode;
                                                        impD_2.DateOB = recipe.Birthday;
                                                        impD_2.RecDate = recipe.RecipeDate;
                                                        impD_2.MNN = recipe.MnnName;
                                                        impD_2.RecNum = recipe.Number;
                                                        impD_2.Quant = recipe.UnitsCount;
                                                        impD_2.MedForm = recipe.MedicationFormName;

                                                        break;

                                                    }
                                                }

                                                //Проверка на наличие рецепта в БД 
                                                var isRecNumDouble2 = await _importDataService.IsRecNumExistAsync(impD_2.RecNum);
                                                if (isRecNumDouble2)
                                                {
                                                        continue;
                                                }
                                                else
                                                {
                                                    importData = impD_2;
                                                    ImportData? _ = await _importDataService.AddEntry(importData);

                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine(ex);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //Waiting period until next cycle (30 second)
                    await Task.Delay(30000);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return listOfMessages;
        }

        

       
        private ImportData getData(Worksheet sheet, int row)
        {
            //Проверка полей с числовым значением и преобразование точек на запятой

            decimal quant = Convert.ToDecimal(sheet.Cell(row, 14).ValueAsString.Replace('.', ','));
            decimal price = Convert.ToDecimal(sheet.Cell(row, 16).ValueAsString.Replace('.', ','));
            decimal pSum = Convert.ToDecimal(sheet.Cell(row, 17).ValueAsString.Replace('.', ','));
            string financeItem = getLetter(sheet.Cell(row, 4).ValueAsString);
            string snils = castSnilsValue(sheet.Cell(row, 22).ValueAsString);
            return new()
            {
                OrgName = sheet.Cell(row, 0).ValueAsString,
                MOD = sheet.Cell(row, 1).ValueAsString,
                INN = sheet.Cell(row, 2).ValueAsString,
                OKPO = sheet.Cell(row, 3).ValueAsString,
                FinancingItem = sheet.Cell(row, 4).ValueAsString,
                ProductName = sheet.Cell(row, 5).ValueAsString,
                MedForm = sheet.Cell(row, 6).ValueAsString,
                SeriaNum = sheet.Cell(row, 7).ValueAsString,
                MNN = sheet.Cell(row, 8).ValueAsString,
                MKB = sheet.Cell(row, 9).ValueAsString,
                RecSeria = sheet.Cell(row, 10).ValueAsString,
                RecNum = sheet.Cell(row, 11).ValueAsString,
                RecDate = sheet.Cell(row, 12).ValueAsDateTime,
                OtpuskDate = sheet.Cell(row, 13).ValueAsDateTime,
                Quant = quant,
                OkeiName = sheet.Cell(row, 15).ValueAsString,
                Price = price,
                PSum = pSum,
                LastName = sheet.Cell(row, 18).ValueAsString,
                Name = sheet.Cell(row, 19).ValueAsString,
                MidName = sheet.Cell(row, 20).ValueAsString,
                DateOB = sheet.Cell(row, 21).ValueAsDateTime,
                SNILS = snils,
                WorkSupplierDogovorId = financeItem,
            };
        }

        public static string StreamToString(Stream stream)
        {
            stream.Position = 0;
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        private ImportDataService _importDataService;
        private ImportDataDuplicateService _doublesService;
      
    }
}
