using Google.Cloud.Firestore;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Final_Project
{
    public partial class Form2 : Form
    {
        Form3 form3 = new Form3();  // 회원가입 폼

        FirestoreDb db; // 파이어 베이스

        internal bool su;       // 권한 여부
        internal string name;   // 이름

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + @"windows-forms-final-project-firebase-adminsdk-pcw1n-5b8352b3ef.json";
            Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", path);

            db = FirestoreDb.Create("windows-forms-final-project");
        }

        /// <summary>
        /// Firebase의 정보와 비교 후 로그인 허용/거부 함수
        /// </summary>
        async void GetMultipleDocumentsFromACollection()
        {   // ID와 PW를 이용한 Hash값 생성
            string hash = form3.SHA256Hash($"{textBox1.Text.Trim()}{textBox2.Text.Trim()}");

            DocumentReference docref = db.Collection("Add_Document_Width_AutoID").Document(hash);
            DocumentSnapshot snap = await docref.GetSnapshotAsync();

            if (snap.Exists)
            {
                Dictionary<string, object> identityDocument = snap.ToDictionary();
                object id;
                object pw;

                identityDocument.TryGetValue("ID", out id);
                identityDocument.TryGetValue("PW", out pw);

                if (id.Equals(textBox1.Text) && pw.Equals(textBox2.Text))
                {
                    Close();
                    DialogResult = DialogResult.OK;
                    su = identityDocument["Su"].Equals("T") ? true : false;
                    name = identityDocument["Name"].ToString();
                }
                else
                {
                    MessageBox.Show("ID 또는 PW가 일치하지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// 회원가입 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            OpenRegisterForm();
        }

        /// <summary>
        /// 로그인 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            GetMultipleDocumentsFromACollection();
        }

        /// <summary>
        /// 회원가입을 위해 Form3을 모달로 여는 함수
        /// </summary>
        public void OpenRegisterForm()
        {
            this.AddOwnedForm(form3);
            form3.ShowDialog(); // 모달 방식으로 form2(로그인 화면)을 실행

            switch (form3.DialogResult)
            {
                case DialogResult.OK:
                    MessageBox.Show("계정 생성을 축하드립니다.", "확인");
                    break;
                case DialogResult.Cancel:
                    MessageBox.Show("회원가입을 취소합니다.", "취소");
                    Close();
                    break;
            }
        }
    }
}
