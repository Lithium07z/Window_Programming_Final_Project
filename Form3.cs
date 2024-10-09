using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Google.Cloud.Firestore;

namespace Final_Project
{
    public partial class Form3 : Form
    {
        FirestoreDb db;

        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + @"windows-forms-final-project-firebase-adminsdk-pcw1n-5b8352b3ef.json";
            Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", path);

            db = FirestoreDb.Create("windows-forms-final-project");
        }

        /// <summary>
        /// SHA256해쉬 함수
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public string SHA256Hash(string data)
        {
            StringBuilder stringBuilder = new StringBuilder();
            SHA256 sha = new SHA256Managed();

            byte[] hash = sha.ComputeHash(Encoding.ASCII.GetBytes(data));
 
            foreach (byte b in hash)
            {
                stringBuilder.AppendFormat("{0:x2}", b);
            }

            return stringBuilder.ToString();
        }

        /// <summary>
        /// 계정 생성 함수
        /// </summary>
        async void DataRegistration()
        {
            string hash = SHA256Hash($"{textBox1.Text.Trim()}{textBox2.Text.Trim()}");

            DocumentReference coll = db.Collection("Add_Document_Width_AutoID").Document(hash);
            Dictionary<string, object> identityDocument = new Dictionary<string, object>()
            {
                { "ID", $"{textBox1.Text}" },
                { "PW", $"{textBox2.Text}" },
                { "Su", $"{domainUpDown1.Text}" },
                { "Name", $"{textBox4.Text}" },
            };

            coll.SetAsync(identityDocument);

            close();
        }

        /// <summary>
        /// 계정 생성 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals(string.Empty) || textBox2.Text.Equals(string.Empty))
            {
                MessageBox.Show("ID와 PW 모두 입력해야 합니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            DataRegistration();

            // 로그인 창의 ID, PW로 전달
            ((Form2)this.Owner).textBox1.Text = this.textBox1.Text;
            ((Form2)this.Owner).textBox2.Text = this.textBox2.Text;

            DialogResult = DialogResult.OK;
        }

        /// <summary>
        /// 취소 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            close();
            DialogResult = DialogResult.Cancel;
        }

        /// <summary>
        /// 필드 초기화 함수
        /// </summary>
        public void close()
        {
            textBox1.Text = textBox2.Text = textBox4.Text = string.Empty;
            Close();
        }

        /// <summary>
        /// 비밀번호 보이기 체크박스 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            textBox2.UseSystemPasswordChar = !textBox2.UseSystemPasswordChar;
        }
    }
}
