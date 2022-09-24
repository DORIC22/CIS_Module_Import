using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
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

namespace CIS_Module_Import
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {       
        IExcelDataReader edr;
        public MainWindow()
        {
            
            InitializeComponent();
            // Добавление выбора проф модуля с БД в ComboBox (ниже)
            using(Model.CISEntities3 db = new Model.CISEntities3()) // используем модель с именем db
            {
                var title = db.ProModule; // назначаем переменной title таблицу Criteria
                foreach (Model.ProModule u in title) // Вроде как цикл нужен чтобы перенести все записи в список
                {
                    string y = u.Title; /* Присваиваем "y - ку", Столбец "Title", 
                                           записm под номером или индексом "u" */
                    CB1.Items.Add(y);   // Добавление переменной y в ComboBox (Имя: CB1).
                }
            }
            // Окончание добавление выбра модуля с БД в ComboBox (выше)
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (CB1.SelectedItem == null)
            {
                MessageBox.Show("Выберите про модуль!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "EXCEL Files (*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() != true)
                    return;

                Dg.ItemsSource = readFile(openFileDialog.FileName);
            }
            
        }

        int IdCriteria_Global;

        private DataView readFile(string fileNames)
        {

            var extension = fileNames.Substring(fileNames.LastIndexOf('.'));
            // Создаем поток для чтения.

            
                FileStream stream = File.Open(fileNames, FileMode.Open, FileAccess.Read);
                // В зависимости от расширения файла Excel, создаем тот или иной читатель.
                // Читатель для файлов с расширением *.xlsx.
                if (extension == ".xlsx")
                    edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
                // Читатель для файлов с расширением *.xls.
                else if (extension == ".xls")
                    edr = ExcelReaderFactory.CreateBinaryReader(stream);

                //// reader.IsFirstRowAsColumnNames
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };
                // Читаем, получаем DataView и работаем с ним как обычно.
                DataSet dataSet = edr.AsDataSet(conf);
                DataView dtView = dataSet.Tables[0].AsDataView();
                DataTable dt = dataSet.Tables[0]; // [0] - выбор листа в Excel
         
            // Модуль поиска и импорта Criteria (ниже)

            int i = 0;
            for (; ; i++)
            {
                if (dt.Rows[i][0].ToString() == "Criteria") // Поиск Criteria
                {                       
                   //MessageBox.Show("Нашел Criteria, строка: " + i); // Просто индикатор, позже удалить.
                    IdCriteria_Global = i+1;
                    break;
                }
            }
            
            i=i + 2; // Для переброса на 2 строчки ниже

            int IdProModule=0;
            for (; ;  i++)
            {
                if (dt.Rows[i][1] != null && dt.Rows[i][1].ToString() != "") /* Двойная проверка на пустые ячейки, 
                                                                              Если ячейка не пустая, производим запись
                                                                            И импорт данных*/
                {
                    string Title = dt.Rows[i][1].ToString();   // Выбор Title (Ячейка "BI" в Excel)
                    string MaxValue = dt.Rows[i][10].ToString(); // Выбор MaxValue (Ячейка "KI" в Excel)

                    Model.CISEntities3 context5 = new Model.CISEntities3();                  
                    IdProModule = context5.ProModule.Single(e => e.Title == CB1.Text).IdProModule;
    
                    var Chck = context5.Criteria.FirstOrDefault(e => e.Title == Title && e.IdProModule == IdProModule);
                    // Проверка на повтор записи
                    
                if (Chck != null)
                {
                        MessageBoxResult result;
                        result = MessageBox.Show("Такие данные уже есть в БД. Произвести их удаление?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Question);
                        switch (result)
                        {
                            case MessageBoxResult.Yes:
                                // Удаление из БД критериев с id выбранным в CB.
                                Model.CISEntities3 db = new Model.CISEntities3();
                                db.Criteria.RemoveRange(db.Criteria.Where(x => x.IdProModule == IdProModule));
                                db.SaveChanges();

                                break;
                            case MessageBoxResult.No:
                                edr.Close();
                                return dtView;
                                break;
                        }
                }
                    
                Model.Criteria TableCriteriaImport = new Model.Criteria(); // 1 Назначение на запись данных из переменных
                TableCriteriaImport.Title = Title; // 2
                TableCriteriaImport.MaxValue = MaxValue; // 3 
                TableCriteriaImport.IdProModule = IdProModule; // 4

                Model.CISEntities3 ModelBase = new Model.CISEntities3(); // Назначение модели
                ModelBase.Criteria.Add(TableCriteriaImport); // Запись
                ModelBase.SaveChanges(); // Сохранение изменений
                //MessageBox.Show(dt.Rows[i][1].ToString()); // Индикатор, удалить на финале
                    
                }
                else
                {
                    MessageBox.Show("Импортированно");
                    break;
                }
            }

            // Окончание модуля поиска и импорта Criteria (выше)
            
            int EndWork2 = 0;

            for (; ;i++ ) // поиск sub criteria (Раздела)
            {
                if (EndWork2 == 1)
                {
                    break;
                }
                if (dt.Rows[i][1].ToString()== "Sub Criterion\nName or Description")
                {
                    IdCriteria_Global++;
                    i++; // Чтобы сразу перейти к SubCriteria

                    for (; ;i++ ) // Цикл по поиску Саб критериев
                    {
                        try
                        {
                            if (dt.Rows[i][1].ToString() != "" && dt.Rows[i][1] != null)
                            { }
                        }
                        catch
                        {
                            MessageBox.Show("Импорт произведен");
                            EndWork2 = 1;
                            break;
                        }
                        if (dt.Rows[i][1].ToString() != "" && dt.Rows[i][1] != null) // Если мы нашлм критерий, то...
                        {
                            int i2 = i+1;
                            int StrokaTitleSubCriteria = i;
                            if (dt.Rows[i][1].ToString() == "Sub Criterion\nName or Description")
                            {
                                i--;
                                break;
                            }
                            double Score = 0; // TotalScoresForAllAspect, пойдет на запись.
                            string SubCriteriaTitle = dt.Rows[i][1].ToString(); //Title, пойдет на запись.


                            i++; // Для перехода к баллам
                           // MessageBox.Show(SubCriteriaTitle.ToString());

                            for (; ;i++ ) // Поиск Баллов
                            {
                                if (dt.Rows[i][10].ToString() != "" && dt.Rows[i][10] != null) // Если ячейка с баллом не пуста, то...
                                {
                                    Score = Score + Convert.ToDouble(dt.Rows[i][10]);
                                   // MessageBox.Show(Score.ToString());
                                }
                                else // Остановка поиска баллов
                                {
                                    break;
                                }
                            } // окончание поиска баллов

                            // Поиск id нужного критерия
                            string RequiredCriteriaText = dt.Rows[IdCriteria_Global][1].ToString(); // записал текст нужного критерия, для поиска его ид в базе
                            Model.CISEntities3 context = new Model.CISEntities3();
                            int IdCriteriaForSubCriteria = context.Criteria.Single(e => e.Title == RequiredCriteriaText && e.IdProModule == IdProModule).IdCriteria;
                            // Конец поиска id Нужного критерия

                            // Запись в БД Сабкритерия
                            Model.SubCriteria SubCriteriaImport = new Model.SubCriteria(); // 1 Назначение на запись данных из переменных
                            SubCriteriaImport.Title = SubCriteriaTitle; // 2
                            SubCriteriaImport.IdCriteria = IdCriteriaForSubCriteria; // 3 
                            SubCriteriaImport.TotalScoresForAllAspects = Score.ToString(); // 4

                            

                            Model.CISEntities3 ModelBase = new Model.CISEntities3(); // Назначение модели
                            ModelBase.SubCriteria.Add(SubCriteriaImport); // Запись
                            try
                            {
                                ModelBase.SaveChanges(); // Сохранение изменений
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.ToString());
                                MessageBox.Show(SubCriteriaTitle, "Title");
                                MessageBox.Show(IdCriteriaForSubCriteria.ToString(), "Id Criteria");
                                MessageBox.Show(Score.ToString(), "Score");
                            }

                            // Запись аспекта

                            for (; ; i2++)
                            {
                                if (dt.Rows[i2][4] == null || dt.Rows[i2][4].ToString() == "")
                                {
                                     break;
                                }
                                else
                                {
                                    string TitleAspect = dt.Rows[i2][4].ToString(); // Title Аспета, на запись
                                    string DescriptionAspect = dt.Rows[i2][6].ToString(); // Description Аспекта, на запись
                                    double ScoreAspect = Convert.ToDouble(dt.Rows[i2][10]); // NumberOfPoints, на запись

                                    string IdTypeAspectExcel = dt.Rows[i2][3].ToString(); // IdTypeAspect, далее будет запрос в бд, чтобы найти id нужного критерия
                                    Model.CISEntities3 context2 = new Model.CISEntities3();
                                    int IdTypeAspect = context2.TypeAspect.Single(e => e.Title == IdTypeAspectExcel).IdTypeAspect; // На запись

                                    string TitleSubCriteriaExcel = dt.Rows[StrokaTitleSubCriteria][1].ToString();
                                    int IdSubCriteria = context2.SubCriteria.Single(e => e.Title == TitleSubCriteriaExcel && e.IdCriteria == IdCriteriaForSubCriteria).IdSubCriteria; // на запись
                                    
                                    if (dt.Rows[i2+1][4].ToString() == "" || dt.Rows[i2+1][4] == null)                              
                                    {
                                        if (dt.Rows[i2 + 1][6].ToString() != "" && dt.Rows[i2 + 1][6] != null)
                                        {
                                            if (IdTypeAspectExcel == "M")
                                            {
                                                int i3 = i2 + 1;
                                                for (; ; i3++)
                                                {
                                                    if (dt.Rows[i3][4] != null && dt.Rows[i3][4].ToString() != "")
                                                    {
                                                        //MessageBox.Show(dt.Rows[i2][4].ToString());
                                                        i2 = i3 - 1;
                                                        break;
                                                    }
                                                    string NewDescriptionAspect = dt.Rows[i3][6].ToString();
                                                    DescriptionAspect = DescriptionAspect + "; " + NewDescriptionAspect;
                                                    //MessageBox.Show(DescriptionAspect);
                                                }
                                            }
                                        }
                                    }

                                    // Запись аспекта в БД

                                    Model.Aspect AspectImport = new Model.Aspect(); // 1 Назначение на запись данных из переменных
                                    AspectImport.Title = TitleAspect;
                                    AspectImport.IdSubCriteria = IdSubCriteria;
                                    AspectImport.NumberOfPoints = ScoreAspect.ToString();
                                    AspectImport.IdTypeAspect = IdTypeAspect;
                                    AspectImport.Description = DescriptionAspect;

                                    Model.CISEntities3 ModelBase2 = new Model.CISEntities3(); // Назначение модели
                                    ModelBase2.Aspect.Add(AspectImport); // Запись
                                   
                                    ModelBase2.SaveChanges(); // Сохранение изменений
                                    

                                    // Конец записи аспекта в БД

                                    // запись в Description of Judje Aspect

                                    if (IdTypeAspectExcel == "J")
                                    {
                                        i2++;
                                        int IdAspect = context2.Aspect.First(e => e.Title == TitleAspect).IdAspect;
                                       
                                        for (; ; i2++ )
                                        {
                                            if (dt.Rows[i2][5] == null || dt.Rows[i2][5].ToString() == "")
                                            {
                                                i2--;
                                                break;
                                            }
                                            int Judg = Convert.ToInt32(dt.Rows[i2][5]);
                                            string DescriptionJudg = dt.Rows[i2][6].ToString();

                                            // запись в бд

                                            Model.DescriptionOfJudgeAspects JudgImport = new Model.DescriptionOfJudgeAspects(); // 1 Назначение на запись данных из переменных
                                            JudgImport.IdAspect = IdAspect;
                                            JudgImport.Judg = Judg.ToString();
                                            JudgImport.Description = DescriptionJudg;

                                            Model.CISEntities3 ModelBase3 = new Model.CISEntities3(); // Назначение модели
                                            ModelBase3.DescriptionOfJudgeAspects.Add(JudgImport); // Запись

                                            ModelBase3.SaveChanges(); // Сохранение изменений
                                        }
            
                                    }


                                }

                            }

                            // Конец записи аспекта


                            // Конец записи в БД сабкритерия
                            i = i - 1;
                            
                        } // Конец поиска саб критерия
                        

                    } // конец цикла по поиску критериев

                } // Ничего


            } // Конец поиска раздела Саб критериев

                // MessageBox.Show(dt.Rows[1][0].ToString()); // [x][y] - x Строка, y - столбец

            // После завершения чтения освобождаем ресурсы.
            edr.Close();
            return dtView;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void Window_LocationChanged(object sender, EventArgs e)
        {
            
        }

        private void Window_LostFocus(object sender, RoutedEventArgs e)
        {
            
        }
    }
}

