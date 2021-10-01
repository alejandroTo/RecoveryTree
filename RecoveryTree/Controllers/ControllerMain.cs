using Microsoft.Office.Interop.Excel;
using RecoveryTree.Models;
using RecoveryTree.Views;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
namespace RecoveryTree.Controllers
{
    class ControllerMain
    {
        Form1 View;
        FormAlert Alert = new FormAlert();
        private int y = 70;
        private int contador = 0;
        private bool flag = true;
        private System.Windows.Forms.TextBox[] Temp;
        private System.Windows.Forms.TextBox[] Temp2;
        public ControllerMain(Form1 view)
        {
            this.View = view;
            view.buttonCreateTextBox.Click += new EventHandler(ButtonCreateTextBox);
            view.buttonCreateTree.Click += new EventHandler(ButtonCreateTree);
        }
        private void ButtonCreateTextBox(object sender, EventArgs e)
        {
            if (flag)
            {
                CreateTxt();
                flag = false;
            }
            else
            {
                CleanTextBox();
                CreateTxt();
            }
            
        }
        public void CreateTxt()
        {
            try
            {
                contador = Convert.ToInt32(View.textBox1.Text.ToString());
                Temp = new System.Windows.Forms.TextBox[contador];
                Temp2 = new System.Windows.Forms.TextBox[contador];
                for (int i = 0; i < contador; i++)
                {
                    Temp[i] = new System.Windows.Forms.TextBox
                    {
                        Height = 100,
                        Width = 70,
                        Location = new System.Drawing.Point(55, y)
                    };
                    Temp[i].Name = "TextBoxInorden" + i;
                    View.Controls.Add(Temp[i]);
                    Temp2[i] = new System.Windows.Forms.TextBox
                    {
                        Height = 100,
                        Width = 70,
                        Location = new System.Drawing.Point(170, y)
                    };
                    Temp2[i].Name = "TextBoxPreorden" + i;
                    View.Controls.Add(Temp2[i]);
                    y += 20;
                }
                y = 70;
            }
            catch (FormatException e)
            {
                Alert.Alert($"Error{e.InnerException}", FormAlert.enmType.Error);
               
            }
            catch (System.OverflowException e)
            {
                Alert.Alert($"Number so Big{e.InnerException}", FormAlert.enmType.Error);
            }
            
        }
        public void CleanTextBox()
        {
            for (int i = 0; i < contador; i++)
            {
                View.Controls.Remove(Temp[i]);
                View.Controls.Remove(Temp2[i]);
            }
        }

        private void ButtonCreateTree(object sender, EventArgs e)
        {
            ExcelMethods ExMeth = new ExcelMethods();
            ExMeth.OpenExcel();
            ExMeth.createSheet("Arbol Examen");
            Queue<string> QuequeInorden = new Queue<string>();
            Queue<string> QuequePreorden = new Queue<string>();
            foreach (var item in Temp)
            {
                QuequeInorden.Enqueue(item.Text);
            }
            foreach (var item in Temp2)
            {
                QuequePreorden.Enqueue(item.Text);
            }
            int size = QuequePreorden.Count;
            Range Celda = Globals.wShe.Cells[1, 1];
            Celda.Value = "-";
            
            for (int i = 1; i <= size; i++)
            {
                ExMeth.FillExcel(1, i + 1, QuequeInorden.Dequeue());
                ExMeth.FillExcel(i + 1, 1, QuequePreorden.Dequeue());
            }
            ExMeth.GetHeaders();
            Alert.Alert("Tree Create Success", FormAlert.enmType.Success);
        }      
    }
}
