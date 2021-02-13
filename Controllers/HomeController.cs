using ExcelProcessor.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Syncfusion.XlsIO;
using System.IO;
using Syncfusion.Drawing;
using System.Reflection;

namespace ExcelProcessor.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            InputModel dummy = new InputModel();
            dummy.auditColumns = new AuditColumns();
            dummy.ProcessResult = new List<ProcessResult>();
            return View(dummy);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult CreateDocument()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult CreateDocument([Bind()] InputModel inputModel)
        {
            string lookupDirectory = @"E:\temp\tvlk\files\";
            string suffix = " - Processed";
            inputModel.ProcessResult = new List<ProcessResult>();

            if (inputModel.Directory != null)
            {
                if (inputModel.Directory.Length > 0)
                {
                    lookupDirectory = inputModel.Directory;
                }
            }

            if (inputModel.OutputSuffix != null)
            {
                if (inputModel.OutputSuffix.Length > 0)
                {
                    suffix = inputModel.OutputSuffix;
                }
            }

            string[] files = Directory.GetFiles(lookupDirectory, "*" + inputModel.SearchString + "*");

            foreach (string file in files)
            {
                if (!file.Contains(suffix))
                {
                    ProcessResult processResult = new ProcessResult();

                    try
                    {
                        //FileStream fileStream = new FileStream(file, FileMode.Open);
                        using (FileStream fileStream = new FileStream(file, FileMode.Open))
                        {
                            //setting up application
                            ExcelEngine excelEngine = new ExcelEngine();
                            IApplication application = excelEngine.Excel;
                            application.DefaultVersion = ExcelVersion.Excel2013;
                            application.EnableIncrementalFormula = true;

                            //setting up workbook & shits
                            IWorkbook workbook = application.Workbooks.Open(fileStream);
                            IWorksheet worksheet = workbook.Worksheets[0];
                            worksheet.EnableSheetCalculations();
                            worksheet.UsedRangeIncludesFormatting = false;

                            //define initial values
                            int colCount = worksheet.UsedRange.LastColumn;
                            int rowCount = worksheet.UsedRange.LastRow;
                            //int colCount = 39;
                            //int rowCount = 10;
                            AuditColumns colName = new AuditColumns();
                            AuditColumns cleanColName = new AuditColumns();
                            AuditColumns colBase = new AuditColumns();
                            ExcelColumnEnum en = new ExcelColumnEnum();

                            //get column alphabet by header text
                            for (int i = 1; i <= colCount; i++)
                            {
                                string[] date = colBase.date.Split(',');
                                foreach (string split in date)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.date = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] bid = colBase.bid.Split(',');
                                foreach (string split in bid)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.bid = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] locale = colBase.locale.Split(',');
                                foreach (string split in locale)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.locale = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] contractEntity = colBase.contractEntity.Split(',');
                                foreach (string split in contractEntity)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.contractEntity = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] contractCurrency = colBase.contractCurrency.Split(',');
                                foreach (string split in contractCurrency)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.contractCurrency = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] collectEntity = colBase.collectEntity.Split(',');
                                foreach (string split in collectEntity)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.collectEntity = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] collectCurrency = colBase.collectCurrency.Split(',');
                                foreach (string split in collectCurrency)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.collectCurrency = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] commissionRevenue = colBase.commissionRevenue.Split(',');
                                foreach (string split in commissionRevenue)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.commissionRevenue = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] transactionFee = colBase.transactionFee.Split(',');
                                foreach (string split in transactionFee)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.transactionFee = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] premium = colBase.premium.Split(',');
                                foreach (string split in premium)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.premium = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] discount = colBase.discount.Split(',');
                                foreach (string split in discount)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.discount = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] coupon = colBase.coupon.Split(',');
                                foreach (string split in coupon)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.coupon = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] redeemedPoints = colBase.redeemedPoints.Split(',');
                                foreach (string split in redeemedPoints)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.redeemedPoints = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] uniqueCode = colBase.uniqueCode.Split(',');
                                foreach (string split in uniqueCode)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.uniqueCode = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] installmentFee = colBase.installmentFee.Split(',');
                                foreach (string split in installmentFee)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.installmentFee = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] deliveryFee = colBase.deliveryFee.Split(',');
                                foreach (string split in deliveryFee)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.deliveryFee = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] invoiceAmount = colBase.invoiceAmount.Split(',');
                                foreach (string split in invoiceAmount)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.invoiceAmount = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] refundFee = colBase.refundFee.Split(',');
                                foreach (string split in refundFee)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.refundFee = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] rescheduleFee = colBase.rescheduleFee.Split(',');
                                foreach (string split in rescheduleFee)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.rescheduleFee = en.TranslateIndex(i);
                                        break;
                                    }
                                }

                                string[] rebookCost = colBase.rebookCost.Split(',');
                                foreach (string split in rebookCost)
                                {
                                    if (worksheet[1, i].Value == split)
                                    {
                                        colName.rebookCost = en.TranslateIndex(i);
                                        break;
                                    }
                                }
                            }

                            cleanColName = cleanUpNonExistingColumn(colName);
                            string colNotFound = getColumnNotFoundWarning(cleanColName);

                            //define new columns header
                            worksheet[1, colCount + 1].Text = "Margin Amount";
                            worksheet[1, colCount + 2].Text = "Margin";
                            worksheet[1, colCount + 3].Text = "Status";

                            //calculate first non-header row & copy to the rest of sheet
                            string formula = getAddingFormula(cleanColName, 2, false) + getSubtractFormula(cleanColName, 2, true);
                            worksheet[2, colCount + 1, rowCount, colCount + 1].Formula = "=" + formula.Replace("#+", "").Replace("#", "");

                            string formula2 = getTaggingResult(2, en.TranslateIndex(colCount + 1));
                            worksheet[2, colCount + 2, rowCount, colCount + 2].Formula = "=" + formula2;

                            worksheet[2, colCount + 3, rowCount, colCount + 3].Text = "ISSUED";

                            //conditional formatting
                            /*
                            //prepare writing to compiled shit
                            List<string> listContractEntity = new List<string>();
                            List<string> listCollectingCurrency = new List<string>();

                            //group contract entity
                            if (cleanColName.contractEntity != "")
                            {
                                var distinct = worksheet[cleanColName.contractEntity + "2:" + cleanColName.contractEntity + rowCount.ToString()].Columns.Distinct();
                                int countContractEntity = distinct.Count();
                                if (countContractEntity > 0)
                                {
                                    foreach (var item in distinct)
                                    {
                                        listContractEntity.Add(item.Value);
                                    }
                                }
                            }

                            //group transaction currency
                            if (cleanColName.collectCurrency != "")
                            {
                                var distinct = worksheet[cleanColName.collectCurrency + "2:" + cleanColName.collectCurrency + rowCount.ToString()].Columns.Distinct();
                                int countCollectCurrency = distinct.Count();
                                if (countCollectCurrency > 0)
                                {
                                    foreach (var item in distinct)
                                    {
                                        listCollectingCurrency.Add(item.Value);
                                    }
                                }
                            }
                            */
                            using (FileStream stream = new FileStream(@"" + file.Split('.')[0] + suffix + "." + file.Split('.')[1], FileMode.Create))
                            {
                                workbook.SaveAs(stream);
                            }

                            workbook.Close();
                            excelEngine.Dispose();

                            processResult.Success = true;
                            processResult.Message = file + " success";

                            if (colNotFound != "")
                            {
                                processResult.Message += " with " + colNotFound + " not found.";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        processResult.Success = false;
                        processResult.Message = file + " failed";
                    }

                    inputModel.ProcessResult.Add(processResult);

                    //write to compiled shit
                }
            }

            ViewData["ProcessResult"] = inputModel.ProcessResult;
            
            return View(inputModel);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public AuditColumns cleanUpNonExistingColumn(AuditColumns input)
        {
            AuditColumns result = new AuditColumns();
            PropertyInfo[] properties = input.GetType().GetProperties();
            foreach (PropertyInfo property in properties)
            {
                string colValue = input.GetType().GetProperty(property.Name).GetValue(input).ToString();
                if (colValue.Length > 2 || colValue.Length <= 0)
                {
                    result.GetType().GetProperty(property.Name).SetValue(result, "");
                }
                else
                {
                    result.GetType().GetProperty(property.Name).SetValue(result, colValue);
                }
            }
            return result;
        }

        public string getColumnNotFoundWarning(AuditColumns input)
        {
            string result = "";
            PropertyInfo[] properties = input.GetType().GetProperties();
            foreach (PropertyInfo property in properties)
            {
                string colValue = input.GetType().GetProperty(property.Name).GetValue(input).ToString();
                if (colValue == "")
                {
                    result += property.Name + ", ";
                }
            }
            return result;
        }

        public string getAddingFormula(AuditColumns input, int initRow, bool negateSign)
        {
            string result = "#";

            if (input.commissionRevenue != "")
                result += "+" + input.commissionRevenue + initRow.ToString();
            if (input.transactionFee != "")
                result += "+" + input.transactionFee + initRow.ToString();
            if (input.premium != "")
                result += "+IF(" + input.premium + initRow.ToString() + ">0," + input.premium + initRow.ToString() + ",0)";
            if (input.refundFee != "")
                result += "+" + input.refundFee + initRow.ToString();
            if (input.rescheduleFee != "")
                result += "+" + input.rescheduleFee + initRow.ToString();
            if (input.rebookCost != "")
                result += "+" + input.rebookCost + initRow.ToString();

            return result;
        }

        public string getSubtractFormula(AuditColumns input, int initRow, bool negateSign)
        {
            string result = "";
            string processOperator = "-";

            if (negateSign)
                processOperator = "+";

            if (input.discount != "")
                result += processOperator + "IF(" + input.discount + initRow.ToString() + "<0," + input.discount + initRow.ToString() + ",0)";
            if (input.coupon != "")
                result += processOperator + input.coupon + initRow.ToString();
            if (input.redeemedPoints != "")
                result += processOperator + input.redeemedPoints + initRow.ToString();
            if (input.uniqueCode != "")
                result += processOperator + input.uniqueCode + initRow.ToString();

            return result;
        }

        public string getTaggingResult(int row, string column)
        {
            string result = "";

            result += "IF(" + column + row.ToString() + "<0,\"negative\",\"positive\")";

            return result;
        }
    }
}
