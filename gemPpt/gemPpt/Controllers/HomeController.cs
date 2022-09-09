using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using gemPpt.Models;
using GemBox.Presentation;
using System.Drawing;
using GemBox.Spreadsheet.Charts;
using GemBox.Spreadsheet;
using ColorName = GemBox.Presentation.ColorName;
using Color = GemBox.Presentation.Color;
using System.Text;
using LengthUnit = GemBox.Presentation.LengthUnit;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Asn1.X500;
using System.Threading;
using RestSharp;
using RestSharp.Authenticators;
using static System.Net.WebRequestMethods;
using Org.BouncyCastle.Utilities;
using System.Xml;
using Newtonsoft.Json;
using Formatting = Newtonsoft.Json.Formatting;
using HarfBuzzSharp;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using static System.Net.Mime.MediaTypeNames;
using Org.BouncyCastle.Asn1.Ocsp;

namespace gemPpt.Controllers;

public class HomeController : Controller
{
  private readonly ILogger<HomeController> _logger;

  public HomeController(ILogger<HomeController> logger)
  {
    _logger = logger;
  }

  public IActionResult Index()
  {

    return View();
  }

  public IActionResult Privacy()
  {

    return View();
  }

  public async Task<ViewResult> Yeni()
  {
    string title, name;

    double x = 1;
    double y = 12;
    double typeX;
    double typeY = 1;
    double numb;
    int sayac = 0;


    string code = "03cbf5ab-1caa-4a44-94b7-2505263aa828";

    ComponentInfo.SetLicense("FREE-LIMITED-KEY");

    SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");



    //using (StreamReader streamReader = new("report.json"))
    //  reportJson = streamReader.ReadToEnd();

    //using (StreamReader streamReader = new("statistics.json"))
    //  statisticsJson = streamReader.ReadToEnd();

    RestClient restClient = new("https://hangwire.kimola.com/v1/Reports");
    RestRequest restRequest = new RestRequest($"/{code}/rows/statistics", Method.Post);

    var key = "QQGDO5mEbjp1yCjWnx9TuA==";

    string json = "{\"content\":null,\"models\":{},\"pieces\":{},\"topics\":[],\"categories\":[],\"from\":0,\"size\":0}";
    restRequest.AddParameter("application/json", json, ParameterType.RequestBody);
    restRequest.AddUrlSegment("code", code);
    restRequest.AddHeader("Authorization", $"Bearer {key}");

    RestResponse<dynamic> statisticsJson = restClient.Execute<dynamic>(restRequest);



    RestClient client = new RestClient("https://hangwire.kimola.com/v1/Reports");
    RestRequest request = new RestRequest($"/{code}", Method.Get);
    request.AddHeader("Authorization", $"Bearer {key}");
    RestResponse<dynamic> reportJson = restClient.Execute<dynamic>(request);

    dynamic report = JsonConvert.DeserializeObject(reportJson.Content);


    dynamic statistics = JsonConvert.DeserializeObject(statisticsJson.Content);



    string reportTitle = report.title ?? report.name;
    int modelCount = statistics.models.Count;

    var presentation = new PresentationDocument();

    // Add new PowerPoint presentation slide.
    var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

    var chart = slide.Content.AddChart(GemBox.Presentation.ChartType.Column,
        5, 5, 326, 114, GemBox.Presentation.LengthUnit.Millimeter);

    // Get underlying Excel chart.
    ExcelChart excelChart = (ExcelChart)chart.ExcelChart;
    ExcelWorksheet worksheet = excelChart.Worksheet;
    int i = 2;
    foreach (var date in statistics.dates)
    {
      string dateTime = date.date;
      int count = date.count;
      dateTime = dateTime.Split(" ")[0];
      worksheet.Cells[$"A{i}"].Value = dateTime;
      worksheet.Cells[$"B{i}"].Value = count;
      i++;
    }
    excelChart.SelectData(worksheet.Cells.GetSubrange($"A{1}:B{i}"), true);

    foreach (var topic in statistics.topics)
    {
      string topicName = topic.name;
      string topicCount = topic.count;
      string topicPercentage = topic.percentage;


      typeX = 2;

      if (topicName.Length >= 5 && topicName.Length < 20)
      {
        numb = Math.Log(topicName.Length, 3.2);
        typeX = typeX * numb;
      }
      else if (topicName.Length >= 20 && topicName.Length < 30)
      {
        numb = Math.Log(topicName.Length, 2.2);
        typeX = typeX * numb;
      }
      else if (topicName.Length >= 30)
      {
        numb = Math.Log(topicName.Length, 2);
        typeX = typeX * numb;
      }


      var sayi = topicName.Length;

      var textBox1 = slide.Content.AddTextBox(
       ShapeGeometryType.RoundedRectangle, x, y, typeX, typeY, LengthUnit.Centimeter);


      //// Set shape format.
      textBox1.Shape.Format.Fill.SetSolid(Color.FromRgb(127, 127, 127));
      textBox1.Shape.Format.Outline.Fill.SetNone();
      textBox1.Shape.Format.Outline.Width = Length.From(1, LengthUnit.Point);

      // Set text box text.
      var text = textBox1.AddParagraph().AddRun(topicName);
      text.Format.Fill.SetSolid(Color.FromRgb(255, 255, 255));
      //// Get text box format.
      var format = textBox1.Format;

      //// Set vertical alignment of the text.
      format.VerticalAlignment = VerticalAlignment.Middle;

      // Set left and top margin.
      format.InternalMarginLeft = Length.From(3, LengthUnit.Millimeter);


      x = x + typeX + 1;

      if (typeX >= 23)
        y = y + typeY + 0.5;

      if (x >= 30.5)
      {
        sayac++;
        if (sayac % 2 != 0)
        {
          //2. satır
          x = 1.5;
        }
        else
        {
          //1.satır
          x = 0.5;

        }

        y = y + typeY + 0.5;
      }

    }

    Dictionary<string, Slide> slides = new Dictionary<string, Slide>();
    var j = 2;

    foreach (var model in statistics.models)
    {
      string modelName = model.name;

      slides.Add($"slide{j}", presentation.Slides.AddNew(SlideLayoutType.Custom));

      //----------TITLE----------
      string title2 = modelName;
      var textBox2 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
          2, 2, 10, 3, GemBox.Presentation.LengthUnit.Centimeter);

      textBox2.AddParagraph().AddRun(title2);
      //graph positions
      double graphPositionX = 2;
      double graphPositionY = 10;
      double graphWidthX = 10;

      //words positions
      double wordPositionX = 2;
      double wordPositionY = 9;

      int count = 1;

      double OrangePercantage;

      foreach (var label in model.statistics)
      {
        if (count == 5)
          break;
        else
        {
          string labelName = label.name;
          string labelCount = label.count;
          string labelPercentage = label.percentage;

          double DlabelPercentage = Convert.ToDouble(labelPercentage); //83
          OrangePercantage = graphWidthX * (DlabelPercentage / 100);
          //graph start
          var shape = slides[$"slide{j}"].Content.AddShape(
              ShapeGeometryType.Rectangle,
                     graphPositionX, graphPositionY, graphWidthX, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

          var format1 = shape.Format;
          var fillFormat1 = format1.Fill;

          shape.Format.Fill.SetSolid(Color.FromRgb(242, 242, 242));
          shape.Format.Outline.Fill.SetNone();

          //graph end

          string label1 = labelName;
          //----------TİTLE sentiment adı----------
          var labelTextBox1 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
              wordPositionX, wordPositionY, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

          var leftLabelTextBox = labelTextBox1.AddParagraph().AddRun(label1);
          leftLabelTextBox.Format.Size = 20;

          //-----------title strong words start

          string label2 = labelCount;
          //----------title weak words start----------
          var labelTextBox2 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
             wordPositionX + 4.2, wordPositionY + 0.15, 3, 1, GemBox.Presentation.LengthUnit.Centimeter);

          var leftLabel1 = labelTextBox2.AddParagraph().AddRun(label2);
          leftLabel1.Format.Fill.SetSolid(Color.FromName(ColorName.Gray));
          leftLabel1.Format.Size = 16;


          //-----------title weak words end

          string label3 = labelPercentage;
          //----------title percentage words start----------
          var labelTextBox3 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
             wordPositionX + 10, wordPositionY + 0.85, 2.5, 0.5, GemBox.Presentation.LengthUnit.Centimeter);

          var boyut = labelTextBox3.AddParagraph().AddRun(label3);
          boyut.Format.Size = 16;


          //-----------title percentage words end.

          //graph start
          //TODO burdaki x degeri saga sola kayacak 
          var shape1 = slides[$"slide{j}"].Content.AddShape(
        ShapeGeometryType.Rectangle,
               graphPositionX, graphPositionY, OrangePercantage, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
          var format2 = shape.Format;
          var fillFormat2 = format2.Fill;

          shape1.Format.Fill.SetSolid(Color.FromRgb(255, 192, 0));
          shape1.Format.Outline.Fill.SetNone();

          //graph end


          //graph positions
          graphPositionX = 2;
          graphPositionY = graphPositionY + 2;

          //words positions
          wordPositionX = 2;
          wordPositionY = wordPositionY + 2;
          count++;

          //--------arada kalan vertical line bölümü başlangıç

          var shapeLine = slides[$"slide{j}"].Content.AddShape(
              ShapeGeometryType.Rectangle,
                     16.5, 9, 0.01, 9, GemBox.Presentation.LengthUnit.Centimeter);

          var formatLine = shapeLine.Format;
          var fillFormatLine = formatLine.Fill;
          shapeLine.Format.Fill.SetSolid(Color.FromRgb(127, 127, 127));
          shapeLine.Format.Fill.SetNone();

          //---------vertical line bölüm bitiş
        }
      }
      j++;
    }


    //      TODO
    //sağ grafik alanı için api hazır değil o yüzden boş

    double x2Degiskeni = 2; // bu da diğer değişkenlerin yuzdelerine bağlanacak.
    int kacGrafikVar = 4; //buna gerek kalmayacak foreach ile dondugumuzde hallolacak


    ////foreach (var item in collection)
    ////{
    ////  ------------sağ label grafik alanı başlangıc
    ////int sağkacGrafikVar = 4;

    ////  graph positions
    ////double rightGraphPositionX = 20;
    ////  double rightGraphPositionY = 10;

    ////  words positions
    ////double rightWordPositionX = 20;
    ////  double rightWordPositionY = 9;

    ////  graph start
    ////  var shapes1 = slides[$"slide{j}"].Content.AddShape(
    ////      ShapeGeometryType.Rectangle,
    ////             rightGraphPositionX, rightGraphPositionY, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    ////  var formats1_1 = shapes1.Format;
    ////  var fillFormats1_1 = formats1_1.Fill;
    ////  shapes1.Format.Fill.SetSolid(Color.FromRgb(127, 127, 127));
    ////  shapes1.Format.Outline.Fill.SetNone();


    ////  graph end

    ////  TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


    ////  string labels1_1 = "sentiment adı ";
    ////  ----------TİTLE sentiment adı----------
    ////  var labelTextBoxs1_1 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
    ////      rightWordPositionX, rightWordPositionY, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    ////  var rightLabelTextboxs1_1 = labelTextBoxs1_1.AddParagraph().AddRun(labels1_1);
    ////  rightLabelTextboxs1_1.Format.Size = 20;

    ////  -----------title strong words start

    ////  string labels2_1 = "silik yazı ";
    ////  ----------title weak words start----------
    ////  var labelTextBoxs2_1 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
    ////      rightWordPositionX + 4.2, rightWordPositionY + 0.15, 3, 1, GemBox.Presentation.LengthUnit.Centimeter);

    ////  var rightLabel1 = labelTextBoxs2_1.AddParagraph().AddRun(labels2_1);
    ////  rightLabel1.Format.Fill.SetSolid(Color.FromName(ColorName.Gray));
    ////  rightLabel1.Format.Size = 16;

    ////  -----------title weak words end

    ////  string labels3_1 = "%43 ";
    ////  ----------title percentage words start----------
    ////  var labelTextBoxs3_1 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
    ////      rightWordPositionX + 10, rightWordPositionY + 0.85, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    ////  var boyuts1 = labelTextBoxs3_1.AddParagraph().AddRun(labels3_1);
    ////  boyuts1.Format.Size = 16;

    ////  -----------title percentage words end

    ////  graph start

    ////  var shapes1_4 = slides[$"slide{j}"].Content.AddShape(
    ////  ShapeGeometryType.Rectangle,
    ////         rightGraphPositionX, rightGraphPositionY, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    ////  var formats1_4 = shapes1_4.Format;
    ////  var fillFormats1_4 = formats1_4.Fill;

    ////  shapes1_4.Format.Fill.SetSolid(Color.FromRgb(155, 187, 89));
    ////  shapes1_4.Format.Outline.Fill.SetNone();

    ////  graph end

    ////  2.grafik başlangıcı
    ////  var shapes1_5 = slides[$"slide{j}"].Content.AddShape(
    //// ShapeGeometryType.Rectangle,
    ////        xDegiskeni + rightWordPositionX, rightGraphPositionY, x2Degiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    ////  var formats1_5 = shapes1_5.Format;
    ////  var fillFormats1_5 = formats1_5.Fill;

    ////  shapes1_5.Format.Fill.SetSolid(Color.FromRgb(149, 55, 53));
    ////  shapes1_5.Format.Outline.Fill.SetNone();


    ////  graph positions
    ////  rightGraphPositionX = 20;
    ////  rightGraphPositionY = rightGraphPositionY + 2;

    ////  words positions
    ////  rightWordPositionX = 20;
    ////  rightWordPositionY = rightWordPositionY + 2;

    ////  2.graph end


    ////-------- - sağ label grafik alanı bitiş
    ////}





    presentation.Save("Created Chart.pptx");
    return View();
  }

  [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
  public IActionResult Error()
  {
    return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
  }
}

