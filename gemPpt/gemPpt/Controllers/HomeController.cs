﻿using System.Diagnostics;
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

    //-----bubble are end-------

    //-----------------------------second page-------

    //TODO: bu alan tekrarlayacak sekilde düzeltilmesi lazım.


    //new slide generated 

    var slide2 = presentation.Slides.AddNew(SlideLayoutType.Custom);

    //new slide generated end


    string title2 = "DEĞİŞİR BURALAR ";
    //----------TITLE----------
    var textBox2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        2, 2, 10, 3, GemBox.Presentation.LengthUnit.Centimeter);

    textBox2.AddParagraph().AddRun(title2);

    //-----------title sonu

    //-----------------left label graphic area started


    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun
    double xDegiskeni = 4;
    double x2Degiskeni = 2;
    int kacGrafikVar = 4;

    //grafik icin pozisyonlar
    double graphPositionX = 2;
    double graphPositionY = 10;

    //kelimeler icin pozisyonlar
    double wordPositionX = 2;
    double wordPositionY = 9;

    //----------------------------------------1. label started ---------------------

    for (int i = 0; i < kacGrafikVar; i++)
    {

      //graph start
      var shape = slide2.Content.AddShape(
          ShapeGeometryType.Rectangle,
                 graphPositionX, graphPositionY, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

      var format1 = shape.Format;
      var fillFormat1 = format1.Fill;

      //graph end

      string label1 = "sentiment adı ";
      //----------TİTLE sentiment adı----------
      var labelTextBox1 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
          wordPositionX, wordPositionY, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

      var leftLabelTextBox = labelTextBox1.AddParagraph().AddRun(label1);
      leftLabelTextBox.Format.Size = 20;

      //-----------title strong words start

      string label2 = "silik yazı ";
      //----------title weak words start----------
      var labelTextBox2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
         wordPositionX + 4.2, wordPositionY + 0.15, 3, 1, GemBox.Presentation.LengthUnit.Centimeter);

      var leftLabel1 = labelTextBox2.AddParagraph().AddRun(label2);
      leftLabel1.Format.Fill.SetSolid(Color.FromName(ColorName.Gray));
      leftLabel1.Format.Size = 16;


      //-----------title weak words end

      string label3 = "%43 ";
      //----------title percentage words start----------
      var labelTextBox3 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
         wordPositionX + 10, wordPositionY +0.85, 2.5, 0.5, GemBox.Presentation.LengthUnit.Centimeter);

      var boyut = labelTextBox3.AddParagraph().AddRun(label3);
      boyut.Format.Size = 16;


      //-----------title percentage words end.

      //graph start

      var shape1 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           graphPositionX, graphPositionY, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
      var format2 = shape.Format;
      var fillFormat2 = format2.Fill;

      shape1.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
      shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

      //graph end


      //grafik icin pozisyonlar
      graphPositionX = 2;
      graphPositionY = graphPositionY + 2;

      //kelimeler icin pozisyonlar
      wordPositionX = 2;
      wordPositionY = wordPositionY + 2;

    }




    //--------arada kalan vertical line bölümü başlangıç

    var shapeLine = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               16.5, 9, 0.01, 9, GemBox.Presentation.LengthUnit.Centimeter);

    var formatLine = shapeLine.Format;
    var fillFormatLine = formatLine.Fill;
    shapeLine.Format.Fill.SetSolid(Color.FromName(ColorName.Gray));
    shapeLine.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));


    //---------vertical line bölüm bitiş




    //------------sağ label grafik alanı başlangıc
    //TODO yorumları sil



    //----------------------------------------1. label başlangıc ---------------------

    //graph start
    var shapes1 = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               20, 10, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var formats1_1 = shapes1.Format;
    var fillFormats1_1 = formats1_1.Fill;

    //graph end

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


    string labels1_1 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBoxs1_1 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        20, 9, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var rightLabelTextboxs1_1 = labelTextBoxs1_1.AddParagraph().AddRun(labels1_1);
    rightLabelTextboxs1_1.Format.Size = 20;

    //-----------title strong words start

    string labels2_1 = "silik yazı ";
    //----------title weak words start----------
    var labelTextBoxs2_1 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        25, 9.15, 3, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var rightLabel1 = labelTextBoxs2_1.AddParagraph().AddRun(labels2_1);
    rightLabel1.Format.Fill.SetSolid(Color.FromName(ColorName.Gray));
    rightLabel1.Format.Size = 16;

    //-----------title weak words end

    string labels3_1 = "%43 ";
    //----------title percentage words start----------
    var labelTextBoxs3_1 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        30, 9.85, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var boyuts1 = labelTextBoxs3_1.AddParagraph().AddRun(labels3_1);
    boyuts1.Format.Size = 16;

    //-----------title percentage words end

    //graph start

    var shapes1_4 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           20, 10, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats1_4 = shapes1_4.Format;
    var fillFormats1_4 = formats1_4.Fill;

    shapes1_4.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shapes1_4.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //graph end

    //2. grafik başlangıcı
    var shapes1_5 = slide2.Content.AddShape(
   ShapeGeometryType.Rectangle,
          xDegiskeni + 20, 10, x2Degiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats1_5 = shapes1_5.Format;
    var fillFormats1_5 = formats1_5.Fill;

    shapes1_5.Format.Fill.SetSolid(Color.FromName(ColorName.Brown));
    shapes1_5.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Brown));


    //2. graph end

    //----------------------------------------1. label bitiş ---------------------


    //----------------------------------------2. label başlangıc ---------------------

    //graph start
    var shapes2_2 = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               20, 12, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var formats2_2 = shapes2_2.Format;
    var fillFormats2_2 = formats2_2.Fill;

    //graph end

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun

    string labels2_2 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBoxs2_2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        20, 11, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var rightLabelTextBoxs2_2 = labelTextBoxs2_2.AddParagraph().AddRun(labels2_2);
    rightLabelTextBoxs2_2.Format.Size = 20;

    //-----------title strong words start

    string labels2_3 = "silik yazı ";
    //----------title weak words start----------
    var labelTextBoxs2_3 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        25, 11.15, 3, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var rightLabel2 = labelTextBoxs2_3.AddParagraph().AddRun(labels2_3);
    rightLabel2.Format.Fill.SetSolid(Color.FromName(ColorName.Gray));
    rightLabel2.Format.Size = 16;
    //-----------title weak words end

    string labels2_4 = "%43 ";
    //----------title percentage words start----------
    var labelTextBoxs2_4 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        30, 11.85, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var boyuts2_4 = labelTextBoxs2_4.AddParagraph().AddRun(labels2_4);
    boyuts2_4.Format.Size = 16;

    //-----------title percentage words end

    //graph start

    var shapes2_5 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           20, 12, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats2_5 = shapes2_5.Format;
    var fillFormats2_5 = formats2_5.Fill;

    shapes2_5.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shapes2_5.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //graph end

    //2. grafik başlangıcı
    var shapes2_6 = slide2.Content.AddShape(
   ShapeGeometryType.Rectangle,
          xDegiskeni + 20, 12, x2Degiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats2_6 = shapes2_6.Format;
    var fillFormats2_6 = formats2_6.Fill;

    shapes2_6.Format.Fill.SetSolid(Color.FromName(ColorName.Brown));
    shapes2_6.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Brown));

    //2. graph end

    //----------------------------------------2. label bitiş ---------------------


    //----------------------------------------3. label başlangıc ---------------------

    //graph start
    var shapes3_1 = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               20, 14, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var formats3_1 = shapes3_1.Format;
    var fillFormats3_1 = formats3_1.Fill;

    //graph end

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


    string labels3_2 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBoxs3_2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        20, 13, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var rightLabelTextBoxs3_2 = labelTextBoxs3_2.AddParagraph().AddRun(labels3_2);
    rightLabelTextBoxs3_2.Format.Size = 20;

    //-----------title strong words start

    string labels3_3 = "silik yazı ";
    //----------title weak words start----------
    var labelTextBoxs3_3 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        25, 13.15, 3, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var rightLabel3 = labelTextBoxs3_3.AddParagraph().AddRun(labels3_3);
    rightLabel3.Format.Fill.SetSolid(Color.FromName(ColorName.Gray));
    rightLabel3.Format.Size = 16;

    //-----------title weak words end

    string labels3_4 = "%43 ";
    //----------title percentage words start----------
    var labelTextBoxs3_4 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        30, 13.85, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var boyuts3_4 = labelTextBoxs3_4.AddParagraph().AddRun(labels3_4);
    boyuts3_4.Format.Size = 16;

    //-----------title percentage words end

    //graph start

    var shapes3_5 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           20, 14, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats3_5 = shapes3_5.Format;
    var fillFormats3_5 = formats3_5.Fill;

    shapes3_5.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shapes3_5.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //graph end

    //2. grafik başlangıcı
    var shapes3_6 = slide2.Content.AddShape(
   ShapeGeometryType.Rectangle,
          xDegiskeni + 20, 14, x2Degiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats3_6 = shapes3_6.Format;
    var fillFormats3_6 = formats3_6.Fill;

    shapes3_6.Format.Fill.SetSolid(Color.FromName(ColorName.Brown));
    shapes3_6.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Brown));


    //2. graph end


    //----------------------------------------3. label bitiş ---------------------


    //----------------------------------------4. label başlangıc ---------------------

    //graph start
    var shapes4_1 = slide2.Content.AddShape(
        ShapeGeometryType.Rectangle,
               20, 16, 10, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var formats4_1 = shapes4_1.Format;
    var fillFormats4_1 = formats4_1.Fill;

    //graph end

    //TODO buradaki x genişliği değişebilir olması lazım ki değere göre değişmiş olsun


    string labels4_2 = "sentiment adı ";
    //----------TİTLE sentiment adı----------
    var labelTextBoxs4_2 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        20, 15, 4.7, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var rightLabelTextBoxs4_2 = labelTextBoxs4_2.AddParagraph().AddRun(labels4_2);
    rightLabelTextBoxs4_2.Format.Size = 20;

    //-----------title strong words start

    string labels4_3 = "silik yazı ";
    //----------title weak words start----------
    var labelTextBoxs4_3 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        25, 15.15, 3, 0.7, GemBox.Presentation.LengthUnit.Centimeter);

    var rightLabel4 = labelTextBoxs4_3.AddParagraph().AddRun(labels4_3);
    rightLabel4.Format.Fill.SetSolid(Color.FromName(ColorName.Gray));
    rightLabel4.Format.Size = 16;

    //-----------title weak words end

    string labels4_4 = "%43 ";
    //----------title percentage words start----------
    var labelTextBoxs4_4 = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle,
        30, 15.85, 2.5, 1, GemBox.Presentation.LengthUnit.Centimeter);

    var boyuts4_4 = labelTextBoxs4_4.AddParagraph().AddRun(labels4_4);
    boyuts4_4.Format.Size = 16;

    //-----------title percentage words end

    //graph start

    var shapes4_5 = slide2.Content.AddShape(
    ShapeGeometryType.Rectangle,
           20, 16, xDegiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats4_5 = shapes4_5.Format;
    var fillFormats4_5 = formats4_5.Fill;

    shapes4_5.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));
    shapes4_5.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

    //graph end
    //2. grafik başlangıcı
    var shapes4_6 = slide2.Content.AddShape(
   ShapeGeometryType.Rectangle,
          xDegiskeni + 20, 16, x2Degiskeni, 0.7, GemBox.Presentation.LengthUnit.Centimeter);
    var formats4_6 = shapes4_6.Format;
    var fillFormats4_6 = formats4_6.Fill;

    shapes4_6.Format.Fill.SetSolid(Color.FromName(ColorName.Brown));
    shapes4_6.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Brown));


    //2. graph end

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

