
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Win32;
using NetOffice;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using OpenQA.Selenium.Chrome;
using SpreadsheetLight;
using System;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using CheckBox = System.Windows.Forms.CheckBox;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        ChromeDriver driver;
        string name;
        string grade;
        string averageGrade;
        string message;
        int nextMessage = 2;
        int lastMessage = 0;
        string fileName;
        string whereToPutTheFiles;
        string evaluationFile;
        int firstOUtOf;
        int lastOutOf;
        string comment = null;
        string comment2 = null;

        List<string> firstSentenceList = new List<string>
            {
                " has a great attitude and a great personality who makes a tremendous effort in class.",
                " is a great student who brings an excellent attitude to class.",
                " is a very clever and a good student.",
                " makes a genuine effort in class and participates well in classroom activities.",
                " is a friendly, conscientious student who always works hard in class and never fails to complete the homework.",
                " is a great child and always comes to the lessons on time and with homework completed.",
                " has made excellent progress throughout the semester.",
                " is an impressive student.",
                " is an excellent example to the other students in behaviour and attitude.",
                " pays attention in class and is well-behaved.",
                " is a friendly student with an excellent attitude to study.",
                " is an excellent student with a great attitude to learning.",
                " is a smart student who tries to know the right answer.",
                " has a lot of potential when it comes to studying English.",
                " is doing very well in class.",
                " comes to class prepared and always completes the homework on time.",
                " is an outstanding student with great knowledge to share in class.",
                " is a fantastic student who enjoys challenges in class.",
                " is a great problem-solver who completes work carefully.",
                " is a very talented child who enjoys learning English.",
                " arrives to class with a happy attitude and is always ready to learn.",
                " is an eager learner and asks questions appropriately when needed.",
                " is very responsible and completes tasks in a timely manner.",
                " has a strong work ethic and thrives well in class.",
                " is very compassionate and is always kind to others.",
                " is helpful and kind and is a pleasure to be around.",
                " is an intelligent student with a respectful attitude towards learning.",
            };

        List<string> secondSentenceList = new List<string>
            {
                " assumes responsibility well and has a fine attitude towards learning.",
                " is very curious and interested in the lessons we learned in class.",
                " is a well behaved child who tries hard to do well in class.",
                " follows most of the class instructions well and doesn't get distracted easily.",
                " is not afraid to speak up in class and tries to get along well with the other students.",
                " is an enthusiastic member of the class who tries to enjoy all aspects of the work.",
                " participates in all activities and always comes to class prepared.",
                " produces great homework and puts a lot of effort into the class.",
                " has a very good attitude towards classwork and homework.",
                " is an intelligent student who makes an effort to participate in class and is well behaved.",
                " always completes the homework and class work thoroughly.",
                " always tries hard in class and completes the homework on time and thoroughly.",
                " finishes the homework on time and tries to make an effort to speak in class.",
                " is very respectful to the teacher and other students.",
                " participates in the class activities and also makes an effort to engage with the teacher.",
                " has settled well into the class and is learning to listen carefully and contribute to the class' lessons.",
                " always does the best and follows the school rules.",
                " is a great role model for other students.",
                " is very cooperative and helpful in class.",
                " has a fantastic attitude towards learning and being a good citizen of the class.",
                " demonstrates positive character traits.",
                " is respectful and considerate towards others.",
                " makes a sincere effort and works hard in class.",
                " is a leader and a positive role model for other students.",
                " has a positive attitude towards school.",
                " is self-confident and has excellent manners.",
                " is a very happy and well-adjusted child.",
                " is a very thoughtful student.",
            };

        List<string> thirdSentenceList = new List<string>
            {
                " has shown a great grasp of the topics we have covered in class.",
                " has demonstrated a full understanding of the topics we have been studying.",
                " has demonstrated an exhaustive understanding of the material we have covered in class.",
                " has shown a good understanding of the topics we have been studying.",
                " has proven to have a good understanding of the topics we have been studying.",
                " has shown an intrinsic understanding of the topics we have been studying.",
                " has proven to have a key grasp of the material we have covered in class.",
                " has exhibited a comprehensive grasp of the material we have been studying.",
                " learns new material well and is making excellent progress.",
                " grasps new material quickly and is making great progress.",
                " rapidly picks up new material and utilises it during the class.",
                " demonstrates an excellent understanding of new material by trying to participate during class.",
                " demonstrates a great comprehension of the materials we are learning in class.",
                " follows the lesson material very well.",
                " masters new material quickly and is making great progress.",
                " did a great job trying to understand the teacher and new material during class.",
                " enjoys participating in conversations and in discussions.",
                " is willing to take part in all classroom activities.",
                " shows initiative and thinks things throughly before responding.",
                " has a well-developed vocabulary and communicates clearly.",
                " exceeds expectations with neat and careful work.",
                " has an impressive understanding and depth of knowledge about the topics.",
                " has strong reading comprehension skills and grasps the concepts well.",
                " is very organized and studies the concepts well.",
                " shows special strengths in English and uses concepts well.",
                " reads smoothly and with good expression.",
                " listens attentively to directions from the teacher and performs tasks well.",
            };
       

        List<string> forthSentenceList = new List<string>
            {
                " is very keen and is progressing rapidly. Keep up the good work!",
                " is such a smart student that can continue to do well.",
                " is doing very well in class.",
                " is achieving great success with English.",
                " is progressing very well.",
                "'s understanding of the lesson material is excellent.",
                " is doing above average.",
                " is achieving good success with English.",
                " grasps new material readily and is making good progress.",
                " has made a huge leap throughout the semester.",
                " is constantly improving and rapid learner.",
                " is very keen and is progressing rapidly. Keep up the good work!",
                " is progressing at a very rapid pace.",
                "'s comprehension of the lesson material is very good.",
                " is such a smart student and will continue to do well in future English classes.",
                " will do well in future English classes.",
                " enjoys being challenged and would be successful in the future.",
                " has learned a great deal this year and has shown great improvements.",
                " is a problem solver and shows a great deal of persistence.",
                " is a very hard-working student and is a joy to teach.",
                " has a positive attitude that will continue to brighten up any classroom.",
                " is a creative student with a bright future.",
                " is doing an excellent job overall this year.",
                " sets high standards and will continue to achieve them endlessly.",
                " is a motivated learner with a future full of success.",
                " excels in English and will continue to do well.",
            };

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            var testing = Microsoft.VisualBasic.Interaction.InputBox("Password", "Attention!", "");

            
            if (testing != "school123")
                System.Environment.Exit(1);


        }

        private void NextStudentButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (nextMessage < lastMessage - 1)
                {
                    nextMessage++;
                    if (comboBox1.SelectedItem != null)
                        using (SLDocument sl = new SLDocument())
                        {
                            /*FileStream fs = new FileStream("D:\\School\\Grades\\2022\\Student Grades - 2nd semester.Cell", FileMode.Open);*/
                            FileStream fs = new FileStream(fileName, FileMode.Open);

                            SLDocument sheet = new SLDocument(fs, comboBox1.SelectedItem.ToString());


                            SLWorksheetStatistics stats = sheet.GetWorksheetStatistics();
                            averageGrade = sheet.GetCellValueAsString(stats.EndRowIndex, 6).Substring(0, 2);


                            name = sheet.GetCellValueAsString(nextMessage, 1);
                            grade = sheet.GetCellValueAsString(nextMessage, 6);
                            message = "< " + comboBox1.SelectedItem.ToString() + "  영어 중간평가 점수안내>" + "\n\n" + name + "\n\n" + "학생점수: " + grade + "점" + "\n" + "반평균: " + averageGrade + "점" + "\n\n" + "시험 치느라 수고 많았습니다. 반 평균보다 점수가 낮은 학생은 기말평가 때 조금 더 노력해주세요                                       .                                                                        ";
                            /*         richTextBox1.Text = richTextBox1.Text + "\n\n" + message;*/
                            richTextBox1.Text = message;

                            
                            /*Clipboard.SetDataObject(message, true, 10, 100);*/
                            firstOUtOf++;
                            OutHowManyStudents.Text = firstOUtOf + "/" + lastOutOf;
                            fs.Close();
                        }
                }
            }
            catch(Exception error)
            {
                ErrorMessageLabel.Text = error.Message;
            }
                
        }

        private void PreviousStudent_Click(object sender, EventArgs e)
        {
            try
            {
                if (nextMessage >= 2)
                {
                    nextMessage--;
                    if (nextMessage != 1)
                    {
                        if (comboBox1.SelectedItem != null)
                            using (SLDocument sl = new SLDocument())
                            {
                                /*FileStream fs = new FileStream("D:\\School\\Grades\\2022\\Student Grades - 2nd semester.Cell", FileMode.Open);*/
                                FileStream fs = new FileStream(fileName, FileMode.Open);

                                SLDocument sheet = new SLDocument(fs, comboBox1.SelectedItem.ToString());


                                SLWorksheetStatistics stats = sheet.GetWorksheetStatistics();
                                averageGrade = sheet.GetCellValueAsString(stats.EndRowIndex, 6).Substring(0, 2);


                                name = sheet.GetCellValueAsString(nextMessage, 1);
                                grade = sheet.GetCellValueAsString(nextMessage, 6);
                                message = "< " + comboBox1.SelectedItem.ToString() + "  영어 중간평가 점수안내>" + "\n\n" + name + "\n\n" + "학생점수: " + grade + "점" + "\n" + "반평균: " + averageGrade + "점" + "\n\n" + "시험 치느라 수고 많았습니다. 반 평균보다 점수가 낮은 학생은 기말평가 때 조금 더 노력해주세요                                       .                                                                        ";
                                /*         richTextBox1.Text = richTextBox1.Text + "\n\n" + message;*/
                                richTextBox1.Text = message;
                                /*Clipboard.SetText(message);*/
                                /*Clipboard.SetDataObject(message, true, 10, 100);*/

                                firstOUtOf--;
                                OutHowManyStudents.Text = firstOUtOf + "/" + lastOutOf;

                                fs.Close();
                            }
                    }
                    else
                    {
                        nextMessage = 2;
                    }
                }
            }
            catch (Exception error)
            {
                ErrorMessageLabel.Text = error.Message;
            }

        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                name = null;
                averageGrade = null;
                message = null;
                nextMessage = 2;
                lastMessage = 0;
                if (comboBox1.SelectedItem != null)
                    using (SLDocument sl = new SLDocument())
                    {
                        /*FileStream fs = new FileStream("D:\\School\\Grades\\2022\\Student Grades - 2nd semester.Cell", FileMode.Open);*/
                        FileStream fs = new FileStream(fileName, FileMode.Open);

                        SLDocument sheet = new SLDocument(fs, comboBox1.SelectedItem.ToString());


                        SLWorksheetStatistics stats = sheet.GetWorksheetStatistics();
                        lastMessage = stats.EndRowIndex;
                        averageGrade = sheet.GetCellValueAsString(stats.EndRowIndex, 6).Substring(0, 2);


                        name = sheet.GetCellValueAsString(nextMessage, 1);
                        grade = sheet.GetCellValueAsString(nextMessage, 6);
                        message = "< " + comboBox1.SelectedItem.ToString() + "  영어 중간평가 점수안내>" + "\n\n" + name + "\n\n" + "학생점수: " + grade + "점" + "\n" + "반평균: " + averageGrade + "점" + "\n\n" + "시험 치느라 수고 많았습니다. 반 평균보다 점수가 낮은 학생은 기말평가 때 조금 더 노력해주세요                                       .                                                                        ";
                        /*         richTextBox1.Text = richTextBox1.Text + "\n\n" + message;*/
                        richTextBox1.Text = message;
                        /*Clipboard.SetText(message);*/
                        /*Clipboard.SetDataObject(message, true, 10, 100);*/

                        firstOUtOf = nextMessage;
                        lastOutOf = lastMessage;

                        firstOUtOf = firstOUtOf - 1;
                        lastOutOf = lastOutOf - 2;
                        OutHowManyStudents.Text = firstOUtOf + "/" + lastOutOf;

                        fs.Close();

                    }
                PreviousStudentButton.Enabled = true;
                NextStudentButton.Enabled = true;
                FindEvaluationFileTemplateButton.Enabled = true;
            }
            catch (Exception error)
            {
                ErrorMessageLabel.Text = error.Message;
            }

        }

        private void GetExelFileButton_Click(object sender, EventArgs e)
        {
            try
            {
                var FD = new System.Windows.Forms.OpenFileDialog();
                FD.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.cell;";
                if (FD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    fileName = FD.FileName;

                    using (SLDocument sl = new SLDocument())
                    {

                        FileStream fs = new FileStream(fileName, FileMode.Open);

                        SLDocument sheet = new SLDocument(fs, "2D");

                        var getAllSheets = sheet.GetSheetNames();
                        foreach (string sheetName in getAllSheets)
                            comboBox1.Items.Add(sheetName);

                        fs.Close();
                    }
                    comboBox1.Enabled = true;
                }
            }
            catch (Exception error)
            {
                ErrorMessageLabel.Text = error.Message;
            }

        }

        private void OutHowManyStudents_Click(object sender, EventArgs e)
        {

        }

        private void CreateFilesForEvaluationButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.Enabled = false;
                string TemplateFileName = evaluationFile;
                


                using (SLDocument sl = new SLDocument())
                {
                    FileStream fs = new FileStream(fileName, FileMode.Open);

                    SLDocument sheet = new SLDocument(fs, comboBox1.SelectedItem.ToString());

                    fs.Close();

                    var testing = sheet.GetSheetNames();

                    SLWorksheetStatistics stats = sheet.GetWorksheetStatistics();
                    averageGrade = sheet.GetCellValueAsString(stats.EndRowIndex, 6);

                    string name = null;
                    string classInfo = null;
                    string schoolName = null;
                    string midtermScore = null;
                    string midtermAverage = null;
                    string finalScore = null;
                    string finalAverage = null;
                    string currentClass = null;
                    string nextLevel = null;
                    string nextClass = null;
                    string gender = null;
                    
                    for (int j = 2; j < stats.EndRowIndex; j++)
                    {
                        
                        name = sheet.GetCellValueAsString(j, 1);
                        classInfo = sheet.GetCellValueAsString(j, 3) + "-" + sheet.GetCellValueAsString(j, 4);
                        schoolName = sheet.GetCellValueAsString(j, 5);
                        midtermScore = sheet.GetCellValueAsString(j, 6);
                        midtermAverage = sheet.GetCellValueAsString(stats.EndRowIndex, 6).Substring(0, 2);
                        finalScore = sheet.GetCellValueAsString(j, 7);
                        finalAverage = sheet.GetCellValueAsString(stats.EndRowIndex, 7).Substring(0, 2);
                        currentClass = sheet.GetCellValueAsString(j, 8);
                        nextLevel = sheet.GetCellValueAsString(j, 9);
                        nextClass = sheet.GetCellValueAsString(j, 10);
                        gender = sheet.GetCellValueAsString(j, 11);

                        var wordApp = new NetOffice.WordApi.Application();
                        var doc = wordApp.Documents.Open(TemplateFileName, false, true);
                        doc.Content.Find.Execute(findText: "{Name}",
                                                    matchCase: false,
                                                    matchWholeWord: false,
                                                    matchWildcards: false,
                                                    matchSoundsLike: false,
                                                    matchAllWordForms: false,
                                                    forward: true, //this may be the one
                                                    wrap: false,
                                                    format: false,
                                                    replaceWith: name,
                                                    replace: WdReplace.wdReplaceAll
                                                    );

                        doc.Content.Find.Execute(findText: "{ClassInfo}",
                                                       matchCase: false,
                                                       matchWholeWord: false,
                                                       matchWildcards: false,
                                                       matchSoundsLike: false,
                                                       matchAllWordForms: false,
                                                       forward: true, //this may be the one
                                                       wrap: false,
                                                       format: false,
                                                       replaceWith: classInfo,
                                                       replace: WdReplace.wdReplaceAll
                                                       );

                        doc.Content.Find.Execute(findText: "{School}",
                                                   matchCase: false,
                                                   matchWholeWord: false,
                                                   matchWildcards: false,
                                                   matchSoundsLike: false,
                                                   matchAllWordForms: false,
                                                   forward: true, //this may be the one
                                                   wrap: false,
                                                   format: false,
                                                   replaceWith: schoolName,
                                                   replace: WdReplace.wdReplaceAll
                                                   );

                        doc.Content.Find.Execute(findText: "{Midterm}",
                                                       matchCase: false,
                                                       matchWholeWord: false,
                                                       matchWildcards: false,
                                                       matchSoundsLike: false,
                                                       matchAllWordForms: false,
                                                       forward: true, //this may be the one
                                                       wrap: false,
                                                       format: false,
                                                       replaceWith: midtermScore,
                                                       replace: WdReplace.wdReplaceAll
                                                       );

                        doc.Content.Find.Execute(findText: "{midAv}",
                                                       matchCase: false,
                                                       matchWholeWord: false,
                                                       matchWildcards: false,
                                                       matchSoundsLike: false,
                                                       matchAllWordForms: false,
                                                       forward: true, //this may be the one
                                                       wrap: false,
                                                       format: false,
                                                       replaceWith: midtermAverage,
                                                       replace: WdReplace.wdReplaceAll
                                                       );

                        doc.Content.Find.Execute(findText: "{final}",
                                                       matchCase: false,
                                                       matchWholeWord: false,
                                                       matchWildcards: false,
                                                       matchSoundsLike: false,
                                                       matchAllWordForms: false,
                                                       forward: true, //this may be the one
                                                       wrap: false,
                                                       format: false,
                                                       replaceWith: finalScore,
                                                       replace: WdReplace.wdReplaceAll
                                                       );

                        doc.Content.Find.Execute(findText: "{finAv}",
                                                       matchCase: false,
                                                       matchWholeWord: false,
                                                       matchWildcards: false,
                                                       matchSoundsLike: false,
                                                       matchAllWordForms: false,
                                                       forward: true, //this may be the one
                                                       wrap: false,
                                                       format: false,
                                                       replaceWith: finalAverage,
                                                       replace: WdReplace.wdReplaceAll
                                                       );

                        doc.Content.Find.Execute(findText: "{CurrentClass}",
                                                      matchCase: false,
                                                      matchWholeWord: false,
                                                      matchWildcards: false,
                                                      matchSoundsLike: false,
                                                      matchAllWordForms: false,
                                                      forward: true, //this may be the one
                                                      wrap: false,
                                                      format: false,
                                                      replaceWith: currentClass,
                                                      replace: WdReplace.wdReplaceAll
                                                      );

                        doc.Content.Find.Execute(findText: "{NextLevel}",
                                                      matchCase: false,
                                                      matchWholeWord: false,
                                                      matchWildcards: false,
                                                      matchSoundsLike: false,
                                                      matchAllWordForms: false,
                                                      forward: true, //this may be the one
                                                      wrap: false,
                                                      format: false,
                                                      replaceWith: nextLevel,
                                                      replace: WdReplace.wdReplaceAll
                                                      );

                        doc.Content.Find.Execute(findText: "{NextClass}",
                                                      matchCase: false,
                                                      matchWholeWord: false,
                                                      matchWildcards: false,
                                                      matchSoundsLike: false,
                                                      matchAllWordForms: false,
                                                      forward: true, //this may be the one
                                                      wrap: false,
                                                      format: false,
                                                      replaceWith: nextClass,
                                                      replace: WdReplace.wdReplaceAll
                                                      );

                         comment = RandomComments(name, gender);
                         comment2 = RandomComments2(name, gender);

                        doc.Content.Find.Execute(findText: "{Comment}",
                                                      matchCase: false,
                                                      matchWholeWord: false,
                                                      matchWildcards: false,
                                                      matchSoundsLike: false,
                                                      matchAllWordForms: false,
                                                      forward: true, //this may be the one
                                                      wrap: false,
                                                      format: false,
                                                      replaceWith: comment,
                                                      replace: WdReplace.wdReplaceAll
                                                      );

                        doc.Content.Find.Execute(findText: "{Comment2}",
                                                      matchCase: false,
                                                      matchWholeWord: false,
                                                      matchWildcards: false,
                                                      matchSoundsLike: false,
                                                      matchAllWordForms: false,
                                                      forward: true, //this may be the one
                                                      wrap: false,
                                                      format: false,
                                                      replaceWith: comment2,
                                                      replace: WdReplace.wdReplaceAll
                                                      );
                        string ResultFileName = whereToPutTheFiles + "\\" + name + ".docx";
                        doc.SaveAs(ResultFileName);
                        doc.Close();
                        wordApp.Quit();
                        
                        if (j >= 30)
                            break;
                    }
                    

                    
                    whereToPutTheFiles = null;
                    evaluationFile = null;
                    CreateFilesForEvaluationButton.Enabled = false;
                    WhereToPutTheFilesButton.Enabled = false;
                    EvaluationFileTemplateName.Text = "Template Name?";
                    FolderLocationName.Text = "Folder Location Name?";
                    this.Enabled = true;





                }
            }
            catch (Exception error)
            {
                whereToPutTheFiles = null;
                evaluationFile = null;
                CreateFilesForEvaluationButton.Enabled = false;
                WhereToPutTheFilesButton.Enabled = false;
                EvaluationFileTemplateName.Text = "Template Name?";
                FolderLocationName.Text = "Folder Location Name?";
                this.Enabled = true;
                ErrorMessageLabel.Text = error.Message;
                
            }

        }

        private void WhereToPutTheFilesButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (var fbd = new FolderBrowserDialog())
                {
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        whereToPutTheFiles = fbd.SelectedPath;
                    }
                }
                CreateFilesForEvaluationButton.Enabled = true;
                FolderLocationName.Text = whereToPutTheFiles;
            }
            catch (Exception error)
            {
                ErrorMessageLabel.Text = error.Message;
            }
        }

        private void FindEvaluationFileTemplateButton_Click(object sender, EventArgs e)
        {
            try
            {
                var FD = new System.Windows.Forms.OpenFileDialog();
                FD.Filter = "Word Files|*.docx;*.doc;";
                if (FD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    evaluationFile = FD.FileName;
                }
                
                WhereToPutTheFilesButton.Enabled = true;
                EvaluationFileTemplateName.Text = evaluationFile;
            }
            catch (Exception error)
            {
                ErrorMessageLabel.Text = error.Message;
            }
        }

        private string RandomComments(string name, string gender)
        {
            string firstComment = null;
            string secondComment = null;

            try
            {
                Random rng = new Random();
                int num = rng.Next(firstSentenceList.Count);
                int num2 = rng.Next(secondSentenceList.Count);


                 firstComment = firstSentenceList[num];
                firstSentenceList.RemoveAt(num);
                 secondComment = secondSentenceList[num2];
                secondSentenceList.RemoveAt(num2);


                
            }
            catch (Exception error)
            {
                ErrorMessageLabel.Text = "here";
            }
            return name + firstComment + " "+ gender + secondComment;

        }

        private string RandomComments2(string name, string gender)
        {
            string thirdComment = null;
            string forthComment = null;
            try
            {
                Random rng = new Random();

                int num3 = rng.Next(thirdSentenceList.Count);
                int num4 = rng.Next(forthSentenceList.Count);


                thirdComment = thirdSentenceList[num3];
                thirdSentenceList.RemoveAt(num3);
                forthComment = forthSentenceList[num4];
                forthSentenceList.RemoveAt(num4);


            }
            catch (Exception error)
            {
                ErrorMessageLabel.Text = "here";
            }
            return gender + thirdComment + " " + name + forthComment;

        }
    }
}