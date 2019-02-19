using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using unvell.ReoGrid;
using unvell.ReoGrid.DataFormat;
using unvell.ReoGrid.Graphics;

namespace Calendar
{
  /// <summary>
  /// Interaktionslogik für MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window
  {
    private Worksheet Worksheet
    {
      get { return grid.CurrentWorksheet; }
    }

    private RangeBorderStyle blackSolid;
    private RangeBorderStyle graySolid;
    private RangeBorderStyle blackBoldSolid;

    private RangeBorderStyle GraySolid
    {
      get
      {
        if (graySolid == null) graySolid = new RangeBorderStyle(SolidColor.Gray, BorderLineStyle.Solid);
        return graySolid;
      }
    }

    private RangeBorderStyle BlackSolid
    {
      get
      {
        if (blackSolid == null) blackSolid = new RangeBorderStyle(SolidColor.Black, BorderLineStyle.Solid);
        return blackSolid;
      }
    }

    private RangeBorderStyle BlackBoldSolid
    {
      get
      {
        if (blackBoldSolid == null) blackBoldSolid = new RangeBorderStyle(SolidColor.Black, BorderLineStyle.BoldSolid);
        return blackBoldSolid;
      }
    }

    public MainWindow()
    {
      InitializeComponent();

      Worksheet.SetSettings(WorksheetSettings.View_ShowHeaders, false);
    }

    private void GenerateHeader(int row, int col)
    {
      Cell c;
      string[] content;

      content = new string[] { "KW", "DT", "Tag", "" };
      for (int i = 0; i < 4; i++)
      {
        Worksheet[row, col] = content[i];
        c = Worksheet.GetCell(row, col);

        c.Border.All = BlackBoldSolid;
        c.Style.HAlign = ReoGridHorAlign.Center;
        c.Style.VAlign = ReoGridVerAlign.Middle;

        col++;
      }
    }

    public static DateTime GetMonday(int week, int year)
    {
      // die 1. KW ist die mit mindestens 4 Tagen im Januar des nächsten Jahres
      DateTime dt = new DateTime(year, 1, 4);

      // Beginn auf Montag setzten
      dt = dt.AddDays(-(int)((dt.DayOfWeek != DayOfWeek.Sunday) ? dt.DayOfWeek - 1 : DayOfWeek.Saturday));

      // Wochen dazu addieren
      return dt.AddDays(--week * 7);
    }

    private void GenerateWeek(int row, int col, int week, int year)
    {
      DateTime day;
      string[] descr = new string[] { "Mo", "Di", "Mi", "Do", "Fr", "Sa", "So" };
      Worksheet.MergeRange(row, col, 5, 1);

      Worksheet.SetRangeBorders(row, col, 5, 3, BorderPositions.All, BlackSolid);
      Worksheet.SetRangeBorders(row, col + 3, 5, 1, BorderPositions.InsideAll, GraySolid);
      Worksheet.SetRangeBorders(row, col, 5, 4, BorderPositions.Outside, BlackBoldSolid);

      Worksheet.SetRangeDataFormat(row, col + 1, 5, 1, CellDataFormatFlag.Text);

      Worksheet.SetRangeStyles(row, col, 5, 4, new WorksheetRangeStyle()
      {
        Flag = PlainStyleFlag.HorizontalAlign | PlainStyleFlag.VerticalAlign,
        HAlign = ReoGridHorAlign.Center,
        VAlign = ReoGridVerAlign.Middle
      }
      );

      Worksheet[row, col] = week;

      day = GetMonday(week, year);

      for (int i = 0; i < 5; i++)
      {
        Worksheet[row + i, col + 1] = day.AddDays(i).ToString("dd.MM");
        Worksheet[row + i, col + 2] = descr[i];
      }
    }

    private void FormatColumns(int col)
    {
      Worksheet.SetColumnsWidth(col, 3, 40);
      Worksheet.SetColumnsWidth(col + 3, 1, 160);
      Worksheet.SetColumnsWidth(col + 4, 1, 20);
    }

    private void GenerateYear(int year)
    {
      int row;
      int col;

      row = 0;
      col = 0;
      for (int week = 1; week <= 52; week++)
      {
        if (row == 0)
        {
          FormatColumns(col);
          GenerateHeader(row, col);
          row++;
        }

        GenerateWeek(row, col, week, year);
        row += 5;

        if (row > 9 * 5)
        {
          row = 0;
          col += 5;
        }
      }
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      GenerateYear(2019);
    }
  }
}
