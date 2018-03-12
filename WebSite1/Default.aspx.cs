using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;


public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            DateTime beginDate = new DateTime();
            DateTime endDate = new DateTime();

            if (DateTime.Now.AddMonths(-1).Month == 12)
            {
                DateTime lM1Date;
                lM1Date = new DateTime(DateTime.Now.Year - 1, 12, 1);
                Calendar1.SelectedDate = lM1Date;
                Calendar1.VisibleDate = DateTime.Now.AddMonths(-1);
                beginDate = Calendar1.SelectedDate;
            }
            else
            //Sets the default begin date (the first date of last month) and the end date (the last date of the last month) on the initial page load
            {
                DateTime lM1Date;
                lM1Date = new DateTime(DateTime.Now.Year, DateTime.Now.Month - 1, 1);
                Calendar1.SelectedDate = lM1Date;
                Calendar1.VisibleDate = DateTime.Now.AddMonths(-1);
                beginDate = Calendar1.SelectedDate;
            }
            
            if (DateTime.Now.AddMonths(-1).Month == 12)
            {
                DateTime llMDate;
                llMDate = new DateTime(DateTime.Now.Year - 1, 12, DateTime.DaysInMonth(DateTime.Now.Year - 1, 12));
                Calendar2.SelectedDate = llMDate;
                Calendar2.VisibleDate = DateTime.Now.AddMonths(-1);
                endDate = Calendar2.SelectedDate;
            }
             else
            {
                DateTime llMDate;
                llMDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month - 1, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month - 1));
                Calendar2.SelectedDate = llMDate;
                Calendar2.VisibleDate = DateTime.Now.AddMonths(-1);
                endDate = Calendar2.SelectedDate;
            }
            
            Label5.Text = "Range is: " + beginDate.ToShortDateString() + " to " + endDate.ToShortDateString();

        }
    }



    protected void Calendar1_SelectionChanged(object sender, EventArgs e)
    {
        //Notifies user of date range being changed
        DateTime beginDate = new DateTime();
        DateTime endDate = new DateTime();
        if (Calendar1.SelectedDate.Month == 1)
        {
            DateTime lM1Date;
            lM1Date = new DateTime(Calendar1.SelectedDate.Year - 1, 12, 1);
            beginDate = lM1Date;
        }
        else

        {
            DateTime lM1Date;
            lM1Date = new DateTime(Calendar1.SelectedDate.Year, Calendar1.SelectedDate.Month - 1, 1);
            beginDate = lM1Date;
        }


        if (Calendar2.SelectedDate.Month == 1)
        {
            DateTime llMDate;
            llMDate = new DateTime(Calendar1.SelectedDate.Year - 1, 12, DateTime.DaysInMonth(DateTime.Now.Year - 1, 12));
            endDate = llMDate;
        }
        else
        {
            DateTime llMDate;
            llMDate = new DateTime(Calendar2.SelectedDate.Year, Calendar2.SelectedDate.Month, DateTime.DaysInMonth(Calendar2.SelectedDate.Year, Calendar2.SelectedDate.Month));
            endDate = llMDate;
        }


        Label5.Text = "Range is: " + beginDate.ToShortDateString() + " to " + endDate.ToShortDateString();
    }

    protected void Calendar2_SelectionChanged(object sender, EventArgs e)
    {        
        //Notifies user of date range being changed
        DateTime beginDate = new DateTime();
        DateTime endDate = new DateTime();

        if (Calendar1.SelectedDate.Month == 1) 
        {
            DateTime lM1Date;
            lM1Date = new DateTime(Calendar2.SelectedDate.Year - 1, 12, 1);
            beginDate = lM1Date;
        }
        else

        {
            DateTime lM1Date;
            lM1Date = new DateTime(Calendar2.SelectedDate.Year, Calendar1.SelectedDate.Month - 1, 1);
            beginDate = lM1Date;
        }

        if (Calendar2.SelectedDate.Month == 1)
        {
            DateTime llMDate;
            llMDate = new DateTime(Calendar2.SelectedDate.Year - 1, 12, DateTime.DaysInMonth(DateTime.Now.Year - 1, 12));
            endDate = llMDate;
        }
        else
        {
            DateTime llMDate;
            llMDate = new DateTime(Calendar2.SelectedDate.Year, Calendar2.SelectedDate.Month - 1, DateTime.DaysInMonth(Calendar2.SelectedDate.Year, Calendar2.SelectedDate.Month - 1));
            endDate = llMDate;
        }
        
        Label5.Text = "Range is: " + beginDate.ToShortDateString() + " to " + endDate.ToShortDateString();
    }

    protected void Button1_OnClick(object sender, EventArgs e)
    {
        DateTime beginDate = new DateTime();
        DateTime endDate = new DateTime();


            if (Calendar1.SelectedDate.Month == 1)
            {
                DateTime lM1Date;
                lM1Date = new DateTime(DateTime.Now.Year - 1, 12, 1);
                beginDate = lM1Date;
            }
            else

            {
                DateTime lM1Date;
                lM1Date = new DateTime(Calendar1.SelectedDate.Year, Calendar1.SelectedDate.Month - 1, 1);
                beginDate = lM1Date;
            }

            if (Calendar2.SelectedDate.Month == 1)
            {
                DateTime llMDate;
                llMDate = new DateTime(DateTime.Now.Year - 1, 12, DateTime.DaysInMonth(DateTime.Now.Year - 1, 12));
                endDate = llMDate;
            }
            else
            {
                DateTime llMDate;
                llMDate = new DateTime(Calendar2.SelectedDate.Year, Calendar2.SelectedDate.Month - 1, DateTime.DaysInMonth(Calendar2.SelectedDate.Year, Calendar2.SelectedDate.Month - 1));
                endDate = llMDate;
            }

        Label5.Text = "Range is: " + beginDate.ToShortDateString() + " to " + endDate.ToShortDateString();


        //Uses DataClassesDataContext from LinqToSql to obtain connection to data from the database
        DataClassesDataContext dtContext = new DataClassesDataContext();

        //LinqToSql query to grab data from the first 15 records
        var query = (from viewApex in dtContext.viewApexes
                     where viewApex.Order_Date >= beginDate && viewApex.Due_Date <= endDate
                     orderby viewApex.Account_Number
                     select new { viewApex.Sold_At, viewApex.Sold_To, viewApex.Account_Number, viewApex.Invoice__, viewApex.Customer_PO__, viewApex.OrderDateText, viewApex.DueDateText, viewApex.Invoice_Total}).Take(15);

        //Gridview to show the user the first 15 records
        GridView1.DataSource = query;
        GridView1.DataBind();

        //Label alerting if no records were found
        if (query.Count() == 0)
            Label6.Visible = true;
        else
            Label6.Visible = false;
    }

protected void Button2_OnClick(object sender, EventArgs e)
    {
        //Uses DataClassesDataContext from LinqToSql to obtain connection to data from the database
        DataClassesDataContext dtContext = new DataClassesDataContext();

        var beginDate = Convert.ToDateTime(Calendar1.SelectedDate.ToShortDateString());
        var endDate = Convert.ToDateTime(Calendar2.SelectedDate.ToShortDateString());

        //LinqToSql query to grab data for all available records
        var query = (from viewApex in dtContext.viewApexes
                         where viewApex.Order_Date >= beginDate && viewApex.Due_Date <= endDate
                     orderby viewApex.Account_Number
                     select new { viewApex.Sold_At, viewApex.Sold_To, viewApex.Account_Number, viewApex.Invoice__, viewApex.Customer_PO__, viewApex.OrderDateText, viewApex.DueDateText, viewApex.Invoice_Total });

        //Creates a list to obtain the records from the LinqToSql query
        List<List<string>> data = new List<List<string>>();
        foreach (dynamic item in query)
        {
            Console.WriteLine(item);
            var Sold_At = (item.Sold_At);
            var Sold_To= (item.Sold_To);
            var Account_Number = (item.Account_Number);
            var Invoice__= (item.Invoice__);
            var Customer_PO__ = (item.Customer_PO__);
            var Order_Date = (item.OrderDateText);
            var Due_Date = (item.DueDateText);
            var Invoice_Total = (item.Invoice_Total);

            data.Add(new List<String> { Sold_At, Sold_To, Account_Number, Invoice__, Customer_PO__, Convert.ToString(Order_Date), Convert.ToString(Due_Date), Convert.ToString(Invoice_Total) });
        }
        query.AsEnumerable();
        
        //Creation of the Excel Package
        ExcelPackage excel = new ExcelPackage();
        var workSheet = excel.Workbook.Worksheets.Add("Output");

        //Defines the column names for the Excel output
        IEnumerable<String> columnnames = new string[] { "Sold At", "Sold To", "Account Number", "Invoice #", "Customer PO #", "Order Date", "Due Date", "Invoice Total" };

        //Utilizes methods from the prvoided SpreadsheetBuilderWithExample.linq (pasted the code into the Class2.cs file)
        SpreadsheetBuilder.SaveData(workSheet, columnnames, data);

        //Process to save the file into an Excel file
        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        Response.AddHeader("content-disposition", "attachment;  filename=Excel.xlsx");
        MemoryStream stream = new MemoryStream(excel.GetAsByteArray());

        Response.OutputStream.Write(stream.ToArray(), 0, stream.ToArray().Length);
        Response.Flush();
        Response.End();
        Response.Close();
    }
}