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
    //// If using Professional version, put your serial key below.
    //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

    //var presentation = PresentationDocument.Load("sample1.pptx");

    //var sb = new StringBuilder();

    //var slide = presentation.Slides[0];

    ////var width = slide.Content.Drawings.Layout.Width;

    //foreach (var shape in slide.Content.Drawings.OfType<Shape>())
    //{
    //  var wit = shape.Name;
    //  for (int i = 0; i < 4; i++)
    //  {
    //    if (shape.Name == ($"bar-{i}"))
    //    {
    //      shape.Layout.Width = 100;
    //    }

    //  }

    //}

    //presentation.Save("1Shape Formatting.pptx");


    return View();
  }

  public IActionResult Privacy()
  {

    ComponentInfo.SetLicense("FREE-LIMITED-KEY");

    var presentation = new PresentationDocument();

    // Create new slide.
    var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);


    // Create new text box.

    double x = 1;
    double y = 1;
    double typeX;
    double typeY = 1.5;
    double numb;
    int sayac = 0;

    List<string> words = new List<string>();
    words.Add("buralara yaz günü kacanım");
    words.Add("as12345678864345653454");
    words.Add("as");
    words.Add("as");
    words.Add("as");
    words.Add("as");
    words.Add("buralara yaz günü kar yağıyor canım");
    words.Add("bornova");
    words.Add("anne");
    words.Add("as");
    words.Add("buralara yaz günü kar yağıyor canım");
    words.Add("allah");
    words.Add("baba");
    words.Add("galaatsaray");
    words.Add("allaaaaaah");
    words.Add("ahsdajshdasd");
    words.Add("buralara yaz günü kar yağıyor canım");
    words.Add("asasdasdas");
    words.Add("asasdasdas");
    words.Add("asasdasdas");
    words.Add(" asdasdfas");


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

      var textBox = slide.Content.AddTextBox(
       ShapeGeometryType.RoundedRectangle, x, y, typeX, typeY, LengthUnit.Centimeter);

      textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

      //// Set shape format.
      textBox.Shape.Format.Fill.SetSolid(Color.FromName(ColorName.LightGray));
      textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));
      textBox.Shape.Format.Outline.Width = Length.From(1, LengthUnit.Point);

      // Set text box text.
      textBox.AddParagraph().AddRun(item);

      //// Get text box format.
      var format = textBox.Format;

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


    presentation.Save("Text Box Formatting.pptx");



    //// If using Professional version, put your serial key below.
    //ComponentInfo.SetLicense("FREE-LIMITED-KEY");

    //var presentation = new PresentationDocument();

    //// Create new slide.
    //var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

    //var textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle,
    //           100, 10, 105, 10, GemBox.Presentation.LengthUnit.Millimeter);

    //textBox.AddParagraph().AddRun("Classification Title");


    //// Create new "rounded rectangle" shape.
    //var shape = slide.Content.AddShape(
    //    ShapeGeometryType.Rectangle,
    //           20, 20, 100, 10, GemBox.Presentation.LengthUnit.Millimeter);
    ////  x eksenindeki yeri  ,  y eksinindeki yeri  ,   x dogrultusunda uzunluk, y doğrultuusndaki genişlik

    //// Get shape format.
    //var format = shape.Format;

    //// Get shape fill format.
    //var fillFormat = format.Fill;

    ////---------------------
    //var presentation1 = new PresentationDocument();

    //// Create new slide.
    //var slide1 = presentation.Slides.AddNew(SlideLayoutType.Custom);

    //// Create new "rounded rectangle" shape.
    //var shape1 = slide.Content.AddShape(
    //    ShapeGeometryType.Rectangle,
    //           20, 20, 50, 10, GemBox.Presentation.LengthUnit.Millimeter);
    ////  x eksenindeki yeri  ,  y eksinindeki yeri  ,   x dogrultusunda uzunluk, y doğrultuusndaki genişlik

    //// Get shape format.
    //var format1 = shape.Format;

    //// Get shape fill format.
    //var fillFormat1 = format.Fill;

    //shape1.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    //shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));


    ////fillFormat.SetSolid(Color.FromName(ColorName.DarkBlue));

    //presentation.Save("Shape Formatting.pptx");


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


    //------------------GRAFİK BÖLÜMÜ BAŞLANGICI--------------

    ComponentInfo.SetLicense("FREE-LIMITED-KEY");

    SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

    var presentation = new PresentationDocument();

    // Add new PowerPoint presentation slide.
    var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

    // Add simple PowerPoint presentation title.
    //TODO title eklenecek.

    // Create PowerPoint chart and add it to slide.
    var chart = slide.Content.AddChart(GemBox.Presentation.ChartType.Column,
        20, 20, 300, 90, GemBox.Presentation.LengthUnit.Millimeter);

    // Get underlying Excel chart.
    ExcelChart excelChart = (ExcelChart)chart.ExcelChart;
    ExcelWorksheet worksheet = excelChart.Worksheet;

    for (int i = 2; i <= grafikSayisi; i++)
    {
      worksheet.Cells[$"A{i}"].Value = words[i - 2];
      worksheet.Cells[$"B{i}"].Value = value[i - 2];
    }

    // Select data.
    excelChart.SelectData(worksheet.Cells.GetSubrange($"A{1}:B{grafikSayisi}"), true);

    //------------GRAFİK BÖLÜMÜ SONU -----------------

    
    //-----------BALONCUKLU ALAN BAŞLANGICI-------------
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

      textBox1.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

      //// Set shape format.
      textBox1.Shape.Format.Fill.SetSolid(Color.FromName(ColorName.LightGray));
      textBox1.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));
      textBox1.Shape.Format.Outline.Width = Length.From(1, LengthUnit.Point);

      // Set text box text.
      textBox1.AddParagraph().AddRun(item);

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

    //-----baloncuklu alan sonu -------

    //ikinci sayfaya geçiyoruz

    //yeni slayt oluştuduk

    var slide2 = presentation.Slides.AddNew(SlideLayoutType.Custom);

    //yeni slayt oluşturma bitti


    //------------------- ortaya atılacak cizgi calısması. ancak dikey yapamadım. TODO

    //// Create new "rounded rectangle" shape.
    //var shape = slide2.Content.AddShape(
    //    ShapeGeometryType.Line,
    //           16.93, 8.53, 10, 0.5, (LengthUnit)VerticalAlignmentStyle.Justify);
    ////  x eksenindeki yeri  ,  y eksinindeki yeri  ,   x dogrultusunda uzunluk, y doğrultuusndaki genişlik

    //// Get shape format.
    //var format1 = shape.Format;

    //// Get shape fill format.
    //var fillFormat1 = format1.Fill;

    //----------------------ortaya atılacak dikey cizgi sonu

    string title2 = "DEĞİŞİR BURALAR ";
    //----------TİTLE----------
    var textBox2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        2, 2, 10, 3, GemBox.Presentation.LengthUnit.Centimeter);

    textBox2.AddParagraph().AddRun(title2);

    //-----------title sonu

    //-----------------sol label grafik alanı başlangıcı



    //----------------------------------------1. label başlangıc ---------------------

    //grafik baslangıcı
    var shape = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               2, 10, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var format1 = shape.Format;
    var fillFormat1 = format1.Fill;

    //grafik sonu

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun
    double xDegiskeni = 4;
    double x2Degiskeni = 2;


    string label1 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBox1 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        2, 9, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBox1.AddParagraph().AddRun(label1);

    //-----------title silik yazı  sonu

    string label2 = "silik yazı ";
    //----------TİTLE silik yazı adı----------
    var labelTextBox2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        6.2, 9, 3, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBox2.AddParagraph().AddRun(label2);

    //-----------title silik yazi  sonu

    string label3 = "%43 ";
    //----------TİTLE silik yazı adı----------
    var labelTextBox3 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        12.5, 9.9, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

   var boyut= labelTextBox3.AddParagraph().AddRun(label3);
    boyut.Format.Size = 16;

    //-----------title silik yazi  sonu

    //grafik baslangıcı

    var shape1 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           2, 10, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var format2 = shape.Format;
    var fillFormat2 = format2.Fill;

    shape1.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //grafik sonu

    //----------------------------------------1. label bitiş ---------------------


    //----------------------------------------2. label başlangıc ---------------------

    //grafik baslangıcı
    var shape2_1 = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               2, 12, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var format2_1 = shape2_1.Format;
    var fillFormat2_1 = format2_1.Fill;

    //grafik sonu

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


    string label2_1 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBox2_1 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        2, 11, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBox2_1.AddParagraph().AddRun(label2_1);

    //-----------title silik yazı  sonu

    string label2_2 = "silik yazı ";
    //----------TİTLE silik yazı adı----------
    var labelTextBox2_2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        6.2, 11, 3, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBox2_2.AddParagraph().AddRun(label2_2);

    //-----------title silik yazi  sonu

    string label2_3 = "%43 ";
    //----------TİTLE silik yazı adı----------
    var labelTextBox2_3 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        12.5, 11.9, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var boyut2_3 = labelTextBox2_3.AddParagraph().AddRun(label2_3);
    boyut2_3.Format.Size = 16;

    //-----------title silik yazi  sonu

    //grafik baslangıcı

    var shape2_4 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           2, 12, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var format2_4 = shape2_4.Format;
    var fillFormat2_4 = format2_4.Fill;

    shape2_4.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shape2_4.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //grafik sonu

    //----------------------------------------2. label bitiş ---------------------


    //----------------------------------------3. label başlangıc ---------------------

    //grafik baslangıcı
    var shape3_1 = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               2, 14, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var format3_1 = shape3_1.Format;
    var fillFormat3_1 = format3_1.Fill;

    //grafik sonu

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


    string label3_1 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBox3_1 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        2, 13, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBox3_1.AddParagraph().AddRun(label3_1);

    //-----------title silik yazı  sonu

    string label3_2 = "silik yazı ";
    //----------TİTLE silik yazı adı----------
    var labelTextBox3_2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        6.2, 13, 3, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBox3_2.AddParagraph().AddRun(label3_2);

    //-----------title silik yazi  sonu

    string label3_3 = "%43 ";
    //----------TİTLE silik yazı adı----------
    var labelTextBox3_3 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        12.5, 13.9, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var boyut3_3 = labelTextBox3_3.AddParagraph().AddRun(label3_3);
    boyut3_3.Format.Size = 16;

    //-----------title silik yazi  sonu

    //grafik baslangıcı

    var shape3_4 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           2, 14, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var format3_4 = shape3_4.Format;
    var fillFormat3_4 = format3_4.Fill;

    shape3_4.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shape3_4.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //grafik sonu

    //----------------------------------------3. label bitiş ---------------------


    //----------------------------------------4. label başlangıc ---------------------

    //grafik baslangıcı
    var shape4_1 = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               2, 16, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var format4_1 = shape4_1.Format;
    var fillFormat4_1 = format4_1.Fill;

    //grafik sonu

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


    string label4_1 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBox4_1 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        2, 15, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBox4_1.AddParagraph().AddRun(label4_1);

    //-----------title silik yazı  sonu

    string label4_2 = "silik yazı ";
    //----------TİTLE silik yazı adı----------
    var labelTextBox4_2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        6.2, 15, 3, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBox4_2.AddParagraph().AddRun(label4_2);

    //-----------title silik yazi  sonu

    string label4_3 = "%43 ";
    //----------TİTLE silik yazı adı----------
    var labelTextBox4_3 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        12.5, 15.9, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var boyut4_3 = labelTextBox4_3.AddParagraph().AddRun(label4_3);
    boyut4_3.Format.Size = 16;

    //-----------title silik yazi  sonu

    //grafik baslangıcı

    var shape4_4 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           2, 16, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var format4_4 = shape.Format;
    var fillFormat4_4 = format4_4.Fill;

    shape4_4.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shape4_4.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //grafik sonu

    //----------------------------------------4. label bitiş ---------------------
    //_-------------------- sol label grafik alanı bitiş






    //------------sağ label grafik alanı başlangıc
    //TODO yorumları sil



    //----------------------------------------1. label başlangıc ---------------------

    //grafik baslangıcı
    var shapes1 = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               20, 10, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var formats1_1 = shapes1.Format;
    var fillFormats1_1 = formats1_1.Fill;

    //grafik sonu

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


    string labels1_1 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBoxs1_1 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        20, 9, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBoxs1_1.AddParagraph().AddRun(labels1_1);

    //-----------title silik yazı  sonu

    string labels2_1 = "silik yazı ";
    //----------TİTLE silik yazı adı----------
    var labelTextBoxs2_1 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        25, 9, 3, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBoxs2_1.AddParagraph().AddRun(labels2_1);

    //-----------title silik yazi  sonu

    string labels3_1 = "%43 ";
    //----------TİTLE silik yazı adı----------
    var labelTextBoxs3_1 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        30, 9.85, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var boyuts1 = labelTextBoxs3_1.AddParagraph().AddRun(labels3_1);
    boyuts1.Format.Size = 16;

    //-----------title silik yazi  sonu

    //grafik baslangıcı

    var shapes1_4 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           20, 10, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats1_4 = shape.Format;
    var fillFormats1_4 = formats1_4.Fill;

    shapes1_4.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shapes1_4.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //grafik sonu

    //2. grafik başlangıcı
    var shapes1_5 = slide2.Content.AddShape(
   ShapeGeometryType.Rectangle,
          xDegiskeni+20, 10, x2Degiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats1_5 = shape.Format;
    var fillFormats1_5 = formats1_5.Fill;

    shapes1_5.Format.Fill.SetSolid(Color.FromName(ColorName.Brown));
    shapes1_5.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Brown));


    //2. grafik sonu

    //----------------------------------------1. label bitiş ---------------------


    //----------------------------------------2. label başlangıc ---------------------

    //grafik baslangıcı
    var shapes2_2 = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               20, 12, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var formats2_2 = shapes2_2.Format;
    var fillFormats2_2 = formats2_2.Fill;

    //grafik sonu

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


    string labels2_2 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBoxs2_2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        20, 11, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBoxs2_2.AddParagraph().AddRun(labels2_2);

    //-----------title silik yazı  sonu

    string labels2_3 = "silik yazı ";
    //----------TİTLE silik yazı adı----------
    var labelTextBoxs2_3 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        25, 11, 3, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBoxs2_3.AddParagraph().AddRun(labels2_3);

    //-----------title silik yazi  sonu

    string labels2_4 = "%43 ";
    //----------TİTLE silik yazı adı----------
    var labelTextBoxs2_4 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        30, 11.85, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var boyuts2_4 = labelTextBoxs2_4.AddParagraph().AddRun(labels2_4);
    boyuts2_4.Format.Size = 16;

    //-----------title silik yazi  sonu

    //grafik baslangıcı

    var shapes2_5 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           20, 12, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats2_5 = shapes2_5.Format;
    var fillFormats2_5 = formats2_5.Fill;

    shapes2_5.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shapes2_5.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //grafik sonu

    //2. grafik başlangıcı
    var shapes2_6 = slide2.Content.AddShape(
   ShapeGeometryType.Rectangle,
          xDegiskeni + 20, 12, x2Degiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats2_6 = shape.Format;
    var fillFormats2_6 = formats2_6.Fill;

    shapes2_6.Format.Fill.SetSolid(Color.FromName(ColorName.Brown));
    shapes2_6.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Brown));

    //2. grafik sonu

    //----------------------------------------2. label bitiş ---------------------


    //----------------------------------------3. label başlangıc ---------------------

    //grafik baslangıcı
    var shapes3_1 = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               20, 14, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var formats3_1 = shapes3_1.Format;
    var fillFormats3_1 = formats3_1.Fill;
    
    //grafik sonu

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


    string labels3_2 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBoxs3_2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        20, 13, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBoxs3_2.AddParagraph().AddRun(labels3_2);

    //-----------title silik yazı  sonu

    string labels3_3 = "silik yazı ";
    //----------TİTLE silik yazı adı----------
    var labelTextBoxs3_3 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        25, 13, 3, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBoxs3_3.AddParagraph().AddRun(labels3_3);

    //-----------title silik yazi  sonu

    string labels3_4 = "%43 ";
    //----------TİTLE silik yazı adı----------
    var labelTextBoxs3_4 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        30, 13.85, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var boyuts3_4 = labelTextBoxs3_4.AddParagraph().AddRun(labels3_4);
    boyuts3_4.Format.Size = 16;

    //-----------title silik yazi  sonu

    //grafik baslangıcı

    var shapes3_5 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           20, 14, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats3_5 = shapes3_5.Format;
    var fillFormats3_5 = formats3_5.Fill;

    shapes3_5.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shapes3_5.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //grafik sonu

    //2. grafik başlangıcı
    var shapes3_6 = slide2.Content.AddShape(
   ShapeGeometryType.Rectangle,
          xDegiskeni + 20, 14, x2Degiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats3_6 = shape.Format;
    var fillFormats3_6 = formats3_6.Fill;

    shapes3_6.Format.Fill.SetSolid(Color.FromName(ColorName.Brown));
    shapes3_6.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Brown));


    //2. grafik sonu


    //----------------------------------------3. label bitiş ---------------------


    //----------------------------------------4. label başlangıc ---------------------

    //grafik baslangıcı
    var shapes4_1 = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               20, 16, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var formats4_1 = shapes4_1.Format;
    var fillFormats4_1 = formats4_1.Fill;

    //grafik sonu

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


    string labels4_2 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBoxs4_2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        20, 15, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBoxs4_2.AddParagraph().AddRun(labels4_2);

    //-----------title silik yazı  sonu

    string labels4_3 = "silik yazı ";
    //----------TİTLE silik yazı adı----------
    var labelTextBoxs4_3 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        25, 15, 3, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    labelTextBoxs4_3.AddParagraph().AddRun(labels4_3);

    //-----------title silik yazi  sonu

    string labels4_4 = "%43 ";
    //----------TİTLE silik yazı adı----------
    var labelTextBoxs4_4 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        30, 15.85, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var boyuts4_4 = labelTextBoxs4_4.AddParagraph().AddRun(labels4_4);
    boyuts4_4.Format.Size = 16;

    //-----------title silik yazi  sonu

    //grafik baslangıcı

    var shapes4_5 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           20, 16, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats4_5 = shapes4_5.Format;
    var fillFormats4_5 = formats4_5.Fill;

    shapes4_5.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shapes4_5.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //grafik sonu
    //2. grafik başlangıcı
    var shapes4_6 = slide2.Content.AddShape(
   ShapeGeometryType.Rectangle,
          xDegiskeni + 20, 16, x2Degiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats4_6 = shape.Format;
    var fillFormats4_6 = formats4_6.Fill;

    shapes4_6.Format.Fill.SetSolid(Color.FromName(ColorName.Brown));
    shapes4_6.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Brown));


    //2. grafik sonu

    //----------------------------------------4. label bitiş ---------------------







    //---------sağ label grafik alanı bitiş





    presentation.Save("Created Chart.pptx");
    return View();
  }

  [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
  public IActionResult Error()
  {
    return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
  }
}

