using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCreator
{
    public class Excel
    {
        public void Gerar(string pathArq, string conteudo, bool formatSheet = false, string sheetName = "")
        {
            Application xlApp;
            Workbook xlBook;
            Worksheet xlSheet1;

            //Iniciar o Excel e criar um novo livro.
            xlApp = new Application();
            xlBook = xlApp.Workbooks.Add(Missing.Value);
            xlSheet1 = (Worksheet)xlBook.Worksheets.get_Item(1);

            //Deleta as Worksheets 3 e 2
            try
            {
                ((Worksheet)xlBook.Worksheets.get_Item(3)).Delete();
                ((Worksheet)xlBook.Worksheets.get_Item(2)).Delete();
            }
            catch (Exception) { }

            //Ativa a Worksheets 1
            xlSheet1.Activate();

            if (!string.IsNullOrEmpty(sheetName))
                xlSheet1.Name = sheetName;

            //Não Apresenta o Excel e não passar o controle para o utilizador.
            xlApp.Visible = false;
            xlApp.UserControl = false;

            //Escreve no arquivo
            int Num_Linha = 0;
            int Num_Coluna = 0;

            string[] strLinhas = conteudo.Split('|');

            foreach (string strL in strLinhas)
            {
                Num_Linha++;
                if (!string.IsNullOrEmpty(strL))
                {
                    string[] strColunas = strL.Split(';');

                    foreach (string strC in strColunas)
                    {
                        Num_Coluna++;
                        xlSheet1.Cells[Num_Linha, Num_Coluna] = strC;

                        if (formatSheet)
                        {
                            if (Num_Linha == 1)
                            {
                                ((Range)xlSheet1.Cells[Num_Linha, Num_Coluna]).Interior.Color = XlRgbColor.rgbDarkGrey;
                                ((Range)xlSheet1.Cells[Num_Linha, Num_Coluna]).Font.Size = 12;
                                ((Range)xlSheet1.Cells[Num_Linha, Num_Coluna]).Font.Color = XlRgbColor.rgbWhite;
                            }
                            else
                            {
                                int valor = 0;
                                if (int.TryParse(strC, out valor) && valor > 0)
                                    ((Range)xlSheet1.Cells[Num_Linha, Num_Coluna]).Interior.Color = XlRgbColor.rgbLightSalmon;
                            }

                            Borders border = ((Range)xlSheet1.Cells[Num_Linha, Num_Coluna]).Borders;
                            border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                            border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                            border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                            border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                        }
                    }

                    Num_Coluna = 0;
                }
            }

            if (formatSheet)
            {
                Range firstLine = (Range)xlSheet1.Cells[1, 1];
                firstLine.EntireRow.Font.Bold = true;


                xlSheet1.Columns.AutoFit();
                xlSheet1.Columns.HorizontalAlignment = XlHAlign.xlHAlignCenter;



            }

            //Fecha o Excel e salva o arquivo
            xlApp.ActiveWindow.Close(true, pathArq, false);
        }
    }
}
