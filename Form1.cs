using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace TextToExcell
{
    public class ImageObject
    {
        public string Type { get; set; }
        public bool IsPerson => Type == "person";
        public bool IsWheelchair => Type == "wheelchair";
    }

    public class ImageObjectCollection
    {
        private List<ImageObject> _imageObjects;
        public List<ImageObject> ImageObjects
        {
            get => (_imageObjects ?? (_imageObjects = new List<ImageObject>()));
            set => _imageObjects = value;
        }

        public string ImageName { get; set; }
        public DateTime Date { get; set; }

        public int PersonCount => ImageObjects.Count(p => p.IsPerson);
        public int VehicleCount => ImageObjects.Count(p => !p.IsPerson && !p.IsWheelchair);
        public int WheelchairCount => ImageObjects.Count(p => p.IsWheelchair);
    }

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var fileLines = System.IO.File.ReadAllLines(@"C:\Users\Mahmood\source\repos\TextToExcell\TextToExcell\uni_all_data.txt");
            var objects = new List<ImageObjectCollection>();
            var index = 0;

            while (index < fileLines.Length)
            {
                if (fileLines[index].StartsWith("Enter Image Path"))
                {
                    var imageInfo = fileLines[index].Split(':')[1];
                    var imageDateParts = imageInfo.Replace(" data/img/Universite_Laval_", "").Replace(".jpg", "").Trim().Split('_');
                    var imageObjectCollection = new ImageObjectCollection()
                    {
                        ImageName = imageInfo.Replace("data/img/", ""),
                        Date = new DateTime(int.Parse(imageDateParts[0]),
                                            int.Parse(imageDateParts[1]),
                                            int.Parse(imageDateParts[2]),
                                            int.Parse(imageDateParts[3]),
                                            int.Parse(imageDateParts[4]),
                                            int.Parse(imageDateParts[5]))
                    };

                    index++;

                    while (index < fileLines.Length && !fileLines[index].StartsWith("Enter Image Path"))
                    {
                        var objectType = fileLines[index].Split(':')[0];
                        if ((new List<string> { "car", "bus", "truck", "person" }).Contains(objectType))
                        {
                            var imageObject = new ImageObject
                            {
                                Type = objectType,
                            };

                            imageObjectCollection.ImageObjects.Add(imageObject);
                        }
                        index++;
                    }
                    objects.Add(imageObjectCollection);
                }
                else
                    index++;
            }

            fileLines = System.IO.File.ReadAllLines(@"C:\Users\Mahmood\source\repos\TextToExcell\TextToExcell\uni_data.txt");
            index = 0;
            while (index < fileLines.Length)
            {
                if (fileLines[index].StartsWith("wheelchair user"))
                {
                    var previousImageInfoLineIndex = index - 1;
                    while (!fileLines[previousImageInfoLineIndex].StartsWith("Enter Image Path"))
                        previousImageInfoLineIndex--;

                    var imageInfo = fileLines[previousImageInfoLineIndex].Split(':')[1];
                    var collection = objects.FirstOrDefault(p => p.ImageName == imageInfo.Replace("data/img/", ""));
                    if (collection == null)
                        throw new Exception("Related image not found");

                    collection.ImageObjects.Add(new ImageObject
                    {
                        Type = "wheelchair"
                    });
                }
                index++;
            }

            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            var worKbooK = excelApp.Workbooks.Add(Type.Missing);
            var worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
            worKsheeT.Name = "Report";
            worKsheeT.Cells[1, 1] = "Image_Name";
            worKsheeT.Cells[1, 2] = "Date";
            worKsheeT.Cells[1, 3] = "Person Count";
            worKsheeT.Cells[1, 4] = "Vehicle Count";
            worKsheeT.Cells[1, 5] = "Wheelchair Count";

            for (var i = 0; i < objects.Count(); i++)
            {
                worKsheeT.Cells[i + 2, 1] = objects[i].ImageName;
                worKsheeT.Cells[i + 2, 2] = objects[i].Date;
                worKsheeT.Cells[i + 2, 3] = objects[i].PersonCount;
                worKsheeT.Cells[i + 2, 4] = objects[i].VehicleCount;
                worKsheeT.Cells[i + 2, 5] = objects[i].WheelchairCount;
            }

            worKbooK.SaveAs(@"C:\Users\Mahmood\source\repos\TextToExcell\TextToExcell\data.xlsx"); ;
            worKbooK.Close();
            excelApp.Quit();
        }
    }
}
