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

  public IActionResult Yeni()
  {
    string title = "";
    string name = "";

    double x = 1;
    double y = 12;
    double typeX;
    double typeY = 1.5;
    double numb;
    int sayac = 0;


    List<string> words = new List<string>();
    words.Add("baba");
    words.Add("galaatsaray");
    words.Add("allaaaaaah");
    words.Add("ahsdajshdasd");
    words.Add("buralara yaz günü kar yağıyor canım");
    words.Add("asasdasdas");
    words.Add("galaatsaray");
    words.Add("allaaaaaah");
    words.Add("ahsdajshdasd");
    words.Add("galaatsaray");
    words.Add("allaaaaaah");
    words.Add("ahsdajshdasd");
    words.Add("galaatsaray");
    words.Add("allaaaaaah");
    words.Add("ahsdajshdasd");

    int grafikSayisi = words.Count;


    List<int> value = new List<int>();
    value.Add(3);
    value.Add(4);
    value.Add(1);
    value.Add(6);
    value.Add(7);
    value.Add(9);
    value.Add(8);
    value.Add(7);
    value.Add(6);
    value.Add(5);
    value.Add(4);
    value.Add(3);
    value.Add(2);
    value.Add(2);
    value.Add(6);


    //------------------graphic part started--------------

    ComponentInfo.SetLicense("FREE-LIMITED-KEY");

    SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

    var presentation = new PresentationDocument();

    // Add new PowerPoint presentation slide.
    var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

    var chart = slide.Content.AddChart(GemBox.Presentation.ChartType.Column,
        20, 20, 300, 90, GemBox.Presentation.LengthUnit.Millimeter);

    // Get underlying Excel chart.
    ExcelChart excelChart = (ExcelChart)chart.ExcelChart;
    ExcelWorksheet worksheet = excelChart.Worksheet;

    for (int i = 2; i <= grafikSayisi; i++)
    {
      worksheet.Cells[$"A{i}"].Value = words[i - 2];
      worksheet.Cells[$"B{i}"].Value = value[i - 2]; //buraya tarihler gelecek. bu yuzden tarihleri farklı bir listede tutmak gerekecek.
    }

    // Select data.
    excelChart.SelectData(worksheet.Cells.GetSubrange($"A{1}:B{grafikSayisi}"), true);

    //------------graphic part end -----------------


    //-----------buble area start-------------
    foreach (var item in words)
    {
      typeX = 2;

      if (item.Length >= 5 && item.Length < 20)
      {
        numb = Math.Log(item.Length, 3.2);
        typeX = typeX * numb;
      }
      else if (item.Length >= 20 && item.Length < 30)
      {
        numb = Math.Log(item.Length, 2.2);
        typeX = typeX * numb;
      }
      else if (item.Length >= 30)
      {
        numb = Math.Log(item.Length, 2);
        typeX = typeX * numb;
      }


      var sayi = item.Length;

      var textBox1 = slide.Content.AddTextBox(
       ShapeGeometryType.RoundedRectangle, x, y, typeX, typeY, LengthUnit.Centimeter);


      //// Set shape format.
      textBox1.Shape.Format.Fill.SetSolid(Color.FromRgb(115, 112, 112));
      textBox1.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Black));
      textBox1.Shape.Format.Outline.Width = Length.From(1, LengthUnit.Point);

      // Set text box text.
      var text = textBox1.AddParagraph().AddRun(item);
      text.Format.Fill.SetSolid(Color.FromRgb(255, 255, 255));
      //// Get text box format.
      var format = textBox1.Format;

      //// Set vertical alignment of the text.
      format.VerticalAlignment = VerticalAlignment.Middle;

      ////// Set left and top margin.
      format.InternalMarginLeft = Length.From(3, LengthUnit.Millimeter);
      //format.InternalMarginTop = Length.From(3, LengthUnit.Millimeter);
      //format.Centered = true;

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

    //-----bubble are end-------

    //-----------------------------second page-------

    //TODO: bu alan tekrarlayacak sekilde düzeltilmesi lazım.



    //-----------title sonu

    //-----------------left label graphic area started


   

    Dictionary<string, Slide> slides = new Dictionary<string, Slide>();

    int modelSayisi = 4; // TODO burası apiden donen isteğe göre SAYFA SAYISINI BELİRLEYECEK


    for (int j = 2; j < modelSayisi + 2; j++)
    {
      //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun
      double xDegiskeni = 4;
      double x2Degiskeni = 2;
      int kacGrafikVar = 4;

      //graph positions
      double graphPositionX = 2;
      double graphPositionY = 10;

      //words positions
      double wordPositionX = 2;
      double wordPositionY = 9;

      //new slide generated 
      slides.Add($"slide{j}", presentation.Slides.AddNew(SlideLayoutType.Custom));


      //new slide generated end


      string title2 = "DEĞİŞİR BURALAR ";
      //----------TITLE----------
      var textBox2 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
          2, 2, 10, 3, GemBox.Presentation.LengthUnit.Centimeter);

      textBox2.AddParagraph().AddRun(title2);


      //----------------------------------------1. label started ---------------------

      for (int i = 0; i < kacGrafikVar; i++)
      {

        //graph start
        var shape = slides[$"slide{j}"].Content.AddShape(
            ShapeGeometryType.Rectangle,
                   graphPositionX, graphPositionY, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

        var format1 = shape.Format;
        var fillFormat1 = format1.Fill;

        shape.Format.Fill.SetSolid(Color.FromRgb(232, 232, 232));
        //shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.LightGray));

        //graph end

        string label1 = "sentiment adı ";
        //----------TİTLE sentiment adı----------
        var labelTextBox1 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
            wordPositionX, wordPositionY, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

        var leftLabelTextBox = labelTextBox1.AddParagraph().AddRun(label1);
        leftLabelTextBox.Format.Size = 20;

        //-----------title strong words start

        string label2 = "silik yazı ";
        //----------title weak words start----------
        var labelTextBox2 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
           wordPositionX + 4.2, wordPositionY + 0.15, 3, 1, GemBox.Presentation.LengthUnit.Centimeter);

        var leftLabel1 = labelTextBox2.AddParagraph().AddRun(label2);
        leftLabel1.Format.Fill.SetSolid(Color.FromName(ColorName.Gray));
        leftLabel1.Format.Size = 16;


        //-----------title weak words end

        string label3 = "%43 ";
        //----------title percentage words start----------
        var labelTextBox3 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
           wordPositionX + 10, wordPositionY + 0.85, 2.5, 0.5, GemBox.Presentation.LengthUnit.Centimeter);

        var boyut = labelTextBox3.AddParagraph().AddRun(label3);
        boyut.Format.Size = 16;


        //-----------title percentage words end.

        //graph start

        var shape1 = slides[$"slide{j}"].Content.AddShape(
      ShapeGeometryType.Rectangle,
             graphPositionX, graphPositionY, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
        var format2 = shape.Format;
        var fillFormat2 = format2.Fill;

        shape1.Format.Fill.SetSolid(Color.FromName(ColorName.Orange));
        //shape1.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Orange));

        //graph end


        //graph positions
        graphPositionX = 2;
        graphPositionY = graphPositionY + 2;

        //words positions
        wordPositionX = 2;
        wordPositionY = wordPositionY + 2;

      }


      //--------arada kalan vertical line bölümü başlangıç

      var shapeLine = slides[$"slide{j}"].Content.AddShape(
          ShapeGeometryType.Rectangle,
                 16.5, 9, 0.01, 9, GemBox.Presentation.LengthUnit.Centimeter);

      var formatLine = shapeLine.Format;
      var fillFormatLine = formatLine.Fill;
      shapeLine.Format.Fill.SetSolid(Color.FromName(ColorName.LightGray));
      //shapeLine.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.LightGray));


      //---------vertical line bölüm bitiş



      //------------sağ label grafik alanı başlangıc
      int sağkacGrafikVar = 4;

      //graph positions
      double rightGraphPositionX = 20;
      double rightGraphPositionY = 10;

      //words positions
      double rightWordPositionX = 20;
      double rightWordPositionY = 9;

      for (int i = 0; i < sağkacGrafikVar; i++)
      {
        //graph start
        var shapes1 = slides[$"slide{j}"].Content.AddShape(
            ShapeGeometryType.Rectangle,
                   rightGraphPositionX, rightGraphPositionY, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

        var formats1_1 = shapes1.Format;
        var fillFormats1_1 = formats1_1.Fill;
        shapes1.Format.Fill.SetSolid(Color.FromRgb(128, 125, 126));


        //graph end

        //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


        string labels1_1 = "sentiment adı ";
        //----------TİTLE sentiment adı----------
        var labelTextBoxs1_1 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
            rightWordPositionX, rightWordPositionY, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

        var rightLabelTextboxs1_1 = labelTextBoxs1_1.AddParagraph().AddRun(labels1_1);
        rightLabelTextboxs1_1.Format.Size = 20;

        //-----------title strong words start

        string labels2_1 = "silik yazı ";
        //----------title weak words start----------
        var labelTextBoxs2_1 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
            rightWordPositionX + 4.2, rightWordPositionY + 0.15, 3, 1, GemBox.Presentation.LengthUnit.Centimeter);

        var rightLabel1 = labelTextBoxs2_1.AddParagraph().AddRun(labels2_1);
        rightLabel1.Format.Fill.SetSolid(Color.FromName(ColorName.Gray));
        rightLabel1.Format.Size = 16;

        //-----------title weak words end

        string labels3_1 = "%43 ";
        //----------title percentage words start----------
        var labelTextBoxs3_1 = slides[$"slide{j}"].Content.AddTextBox(ShapeGeometryType.Rectangle,
            rightWordPositionX + 10, rightWordPositionY + 0.85, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

        var boyuts1 = labelTextBoxs3_1.AddParagraph().AddRun(labels3_1);
        boyuts1.Format.Size = 16;

        //-----------title percentage words end

        //graph start

        var shapes1_4 = slides[$"slide{j}"].Content.AddShape(
        ShapeGeometryType.Rectangle,
               rightGraphPositionX, rightGraphPositionY, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
        var formats1_4 = shapes1_4.Format;
        var fillFormats1_4 = formats1_4.Fill;

        shapes1_4.Format.Fill.SetSolid(Color.FromRgb(155, 181, 40));
        //shapes1_4.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Orange));

        //graph end

        //2. grafik başlangıcı
        var shapes1_5 = slides[$"slide{j}"].Content.AddShape(
       ShapeGeometryType.Rectangle,
              xDegiskeni + rightWordPositionX, rightGraphPositionY, x2Degiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
        var formats1_5 = shapes1_5.Format;
        var fillFormats1_5 = formats1_5.Fill;

        shapes1_5.Format.Fill.SetSolid(Color.FromName(ColorName.Brown));
        //shapes1_5.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Brown));


        //graph positions
        rightGraphPositionX = 20;
        rightGraphPositionY = rightGraphPositionY + 2;

        //words positions
        rightWordPositionX = 20;
        rightWordPositionY = rightWordPositionY + 2;

        //2. graph end
      }

      //---------sağ label grafik alanı bitiş

    }





    presentation.Save("Created Chart.pptx");
    return View();
  }

  [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
  public IActionResult Error()
  {
    return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
  }
}

