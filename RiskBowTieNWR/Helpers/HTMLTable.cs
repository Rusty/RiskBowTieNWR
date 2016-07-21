using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RiskBowTieNWR.Helpers
{
    public class HTMLTable
    {
        private int Rows { get; set; }
        private int Cols { get; set; }
        private string[] Data { get; set; }
        private string[] Colors { get; set; }

        private double[] ColWidths { get; set; }

        private void EnsureSpace()
        {
            Data = new string[Rows * Cols];
            Colors = new string[Rows * Cols];
            ColWidths = new double[Cols];
            for (int i = 0; i < Cols; i++)
                ColWidths[i] = 100;
        }

        public HTMLTable(int cols)
        {
            Rows = 1;
            Cols = cols;
            EnsureSpace();

        }
        public HTMLTable(int rows, int cols)
        {
            Rows = rows;
            Cols = cols;
            EnsureSpace();
        }

        public bool SetValue(int row, int col, string text, string color=null)
        {
            //Debug.WriteLine($"test='{text}' Colour='{color}'");

            if (row >= Rows)
            {
                var origData = Data;
                var origColors = Colors;
                Rows = row + 1;
                EnsureSpace();
                for (int i = 0; i < origData.Length; i++)
                {
                    Data[i] = origData[i];
                    Colors[i] = origColors[i];
                }
            }

            if (row >= 0 && row < Rows)
                if (col >= 0 && col < Cols)
                {
                    int index = row * Cols + col;

                    Data[index] = text;
                    if (color != null)
                        Colors[index] = color;
                    return true;
                }

            return false;
        }

        private string _tableWidth = "100%";

        public bool SetTableWidthPercentage(double percentage)
        {
            if (percentage > 0)
            {
                _tableWidth = percentage + "%";
                return true;
            }
            return false;
        }


        public bool SetColWidth(int col, double percentage)
        {
            if (col >= 0 && col < Cols)
            {
                ColWidths[col] = percentage;
                return true;
            }
            return false;
        }

        public string GetHTML
        {
            get
            {
                var widthPercent = 100 / Cols;

                string data = @"
<html><head>
<style type=""text/css"">
.c0 { font-family: '/SC.RoadmapViewControl;component/Fonts/Fonts.zip#Museo Sans 300'; font-size: 8px} 
.c1 { border-collapse: collapse; width: tablewidthpercentage } 
.c2 { width: {columnWidth}% } 
.c3 { border-color: Black; border-style: solid; border-width: thin; padding: 0px 7px } 
.c4 { margin: 0px } 
</style>
</head>
<body class=""c0"">
<table class=""c1"">";

                data = data.Replace("tablewidthpercentage", _tableWidth);


                // add the columns
                for (int c = 0; c < Cols; c++)
                    data += string.Format("<col class=\"c2\" width={0}%/>", ColWidths[c]);
                // finish the table heading
                data += "</tr>";

                // add the cell data
                for (int r = 0; r < Rows; r++)
                {
                    for (int c = 0; c < Cols; c++)
                        data += GetCellHTML(r, c);
                    data += "</tr>";
                }

                data += "</table></body></html>";
                data = data.Replace("{columnWidth}", widthPercent.ToString());

                return data;
            }
        }

        private string GetCellHTML(int row, int col)
        {
            int index = row*Cols + col;
            if (!string.IsNullOrWhiteSpace(Colors[index]))
                return string.Format($"<td bgcolor=\"{Colors[index]}\" class=\"c3\"><p class=\"c4\">{Data[index]}</p></td>");

            return string.Format($"<td class=\"c3\"><p class=\"c4\">{Data[index]}</p></td>");
        }
    }
}
