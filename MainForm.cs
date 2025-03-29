using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AibaseResumeScreening
{
    using System;
    using System.Drawing;
    using System.Windows.Forms;
    using System.IO;
    using System.Linq;
    using System.Text;
    using iText.Kernel.Pdf;
    using iText.Kernel.Pdf.Canvas.Parser;
    using Microsoft.Office.Interop.Word;
    using System.Data.SqlClient;
    using System.Configuration;

    namespace AIResumeScreeningSystem
    {
        public class MainForm : Form
        {
            private Button btnUploadResume;
            private Button btnAnalyzeResume;
            private TextBox txtJobDescription;
            private DataGridView dataGridViewCandidates;
            private Label lblStatus;
            private string uploadedResumePath = "";

            public MainForm()
            {
                InitializeUI();
            }

            private void InitializeUI()
            {
                // Form Properties
                this.Text = "AI-Based Resume Screening System";
                this.Size = new Size(800, 500);
                this.StartPosition = FormStartPosition.CenterScreen;

                // Upload Resume Button
                btnUploadResume = new Button();
                btnUploadResume.Text = "Upload Resume";
                btnUploadResume.Size = new Size(150, 30);
                btnUploadResume.Location = new Point(50, 30);
                btnUploadResume.Click += BtnUploadResume_Click;
                this.Controls.Add(btnUploadResume);

                // Job Description TextBox
                txtJobDescription = new TextBox();
                txtJobDescription.Multiline = true;
                txtJobDescription.Size = new Size(600, 100);
                txtJobDescription.Location = new Point(50, 80);
                txtJobDescription.Text = "Enter Job Description...";
                this.Controls.Add(txtJobDescription);

                // Analyze Resume Button
                btnAnalyzeResume = new Button();
                btnAnalyzeResume.Text = "Analyze Resume";
                btnAnalyzeResume.Size = new Size(150, 30);
                btnAnalyzeResume.Location = new Point(50, 200);
                btnAnalyzeResume.Click += BtnAnalyzeResume_Click;
                this.Controls.Add(btnAnalyzeResume);

                // Candidates DataGridView
                dataGridViewCandidates = new DataGridView();
                dataGridViewCandidates.Size = new Size(600, 150);
                dataGridViewCandidates.Location = new Point(50, 250);
                dataGridViewCandidates.ColumnCount = 3;
                dataGridViewCandidates.Columns[0].Name = "Candidate Name";
                dataGridViewCandidates.Columns[1].Name = "Match Score";
                dataGridViewCandidates.Columns[2].Name = "Remarks";
                this.Controls.Add(dataGridViewCandidates);

                // Status Label
                lblStatus = new Label();
                lblStatus.Text = "Status: Ready";
                lblStatus.Size = new Size(200, 30);
                lblStatus.Location = new Point(50, 420);
                this.Controls.Add(lblStatus);
            }

            private void BtnUploadResume_Click(object sender, EventArgs e)
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "PDF Files|*.pdf|Word Documents|*.docx";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    uploadedResumePath = openFileDialog.FileName;
                    lblStatus.Text = "Resume Uploaded: " + uploadedResumePath;

                    MessageBox.Show("Resume Uploaded Successfully!");
                }
            }

            private void BtnAnalyzeResume_Click(object sender, EventArgs e)
            {
                if (string.IsNullOrEmpty(uploadedResumePath))
                {
                    MessageBox.Show("Please upload a resume first.");
                    return;
                }

                lblStatus.Text = "Analyzing resume...";

                string jobDescription = txtJobDescription.Text;
                string resumeText = ExtractTextFromResume(uploadedResumePath);

                double matchScore = CalculateMatchScore(jobDescription, resumeText);
                string remarks = matchScore > 50 ? "Good Match" : "Needs Review";

                // Display in DataGridView
                dataGridViewCandidates.Rows.Add("Candidate 1", $"{matchScore:F2}%", remarks);

                // Save to Database
                SaveToDatabase("Candidate 1", matchScore, remarks);

                lblStatus.Text = "Analysis Complete!";
                MessageBox.Show("Resume Analysis Complete!");
            }

            private string ExtractTextFromResume(string filePath)
            {
                if (filePath.EndsWith(".pdf"))
                    return ExtractTextFromPDF(filePath);
                else if (filePath.EndsWith(".docx"))
                    return ExtractTextFromDOCX(filePath);

                return "";
            }

            private string ExtractTextFromPDF(string filePath)
            {
                StringBuilder text = new StringBuilder();
                using (PdfReader reader = new PdfReader(filePath))
                using (PdfDocument pdfDoc = new PdfDocument(reader))
                {
                    for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                    {
                        text.Append(PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i)));
                    }
                }
                return text.ToString();
            }

            private string ExtractTextFromDOCX(string filePath)
            {
                Application wordApp = new Application();
                Document doc = wordApp.Documents.Open(filePath);
                string text = doc.Content.Text;
                doc.Close();
                wordApp.Quit();
                return text;
            }

            private double CalculateMatchScore(string jobDescription, string resumeText)
            {
                string[] jobKeywords = jobDescription.Split(new char[] { ' ', ',', '.', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                string[] resumeWords = resumeText.Split(new char[] { ' ', ',', '.', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                int matchCount = jobKeywords.Count(keyword => resumeWords.Contains(keyword, StringComparer.OrdinalIgnoreCase));
                return ((double)matchCount / jobKeywords.Length) * 100;
            }

            private void SaveToDatabase(string candidateName, double matchScore, string remarks)
            {
                string connectionString = ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString;

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string query = "INSERT INTO ResumeAnalysis (CandidateName, MatchScore, Remarks) VALUES (@name, @score, @remarks)";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@name", candidateName);
                        cmd.Parameters.AddWithValue("@score", matchScore);
                        cmd.Parameters.AddWithValue("@remarks", remarks);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            [STAThread]
            public static void Main()
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
        }
    }

}
