using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Google.Cloud.Firestore;
using System.Drawing.Printing;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;

namespace Final_Project
{
    // 스크린샷 캡쳐를 위한 Rect 구조체 정의
    [StructLayout(LayoutKind.Sequential)]
    public struct Rect
    {   // 창의 좌, 상, 우, 하 좌표
        public int left;
        public int top;
        public int right;
        public int bottom;
    }

    public partial class Form1 : Form
    {
        Form2 form2 = new Form2();  // 로그인 폼
        Form3 form3 = new Form3();  // 회원가입 폼

        FirestoreDb db; // 파이어 베이스

        private List<string> images = new List<string>();   // 이미지 경로 리스트

        private string pageSet; // 인쇄 설정 저장

        private ListView listView;

        public Form1()
        {
            InitializeComponent();  // 컴포넌트 초기화
            InitializeInStockDate();// 재고 페이지에 현재 날짜 기준으로 입고 정보 추가
            LoadImages();           // 홈 페이지 이미지 로드
            ShowCurrentImage();     // 최초 로드 시 첫번째 이미지를 보여줌
            pageSet = string.Empty; // 인쇄 설정 초기화
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Firebase 코드
            string path = AppDomain.CurrentDomain.BaseDirectory + @"deleted";
            Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", path);

            db = FirestoreDb.Create("windows-forms-final-project");

            // 로그인 코드
            OpenLoginForm();    // 폼이 로드되기 직전에 먼저 로그인 폼부터 모달로 띄우고 확인

            // 계정 페이지 불러오기 코드
            GetALLDocumentsFromACollection();

            // 홈 페이지 시계 코드
            timer1.Interval = 100;  // 타이머 간격 100ms
            timer1.Start();         // 타이머 시작

            // 홈 페이지 인사말 설정
            label15.Text = $"안녕하세요. {form2.name}님";
        }

        /***********************************************************************
        *                             보조, 로직 함수                            
        ***********************************************************************/
        #region .

        /// <summary>
        /// 로그인 폼 로그인, 종료 함수
        /// </summary>
        public void OpenLoginForm()
        {
            this.AddOwnedForm(form2);
            form2.ShowDialog(); // 모달 방식으로 form2(로그인 화면)을 실행
            switch (form2.DialogResult) // 모달 결과 확인
            {
                case DialogResult.OK:
                    MessageBox.Show("로그인되었습니다.", "로그인");
                    break;
                case DialogResult.Cancel:
                    MessageBox.Show("프로그램이 종료됩니다.", "종료");
                    Close();
                    break;
            }
        }

        #region 캡쳐

        // WinAPI 함수를 사용하기 위해 DLL을 임포트
        [DllImport("user32.dll")]
        private static extern int SetForegroundWindow(IntPtr hWnd); // 특정 창을 포커스로 설정

        private const int SW_RESTORE = 9;   // 창을 복원 상태로 만드는 상수

        [DllImport("user32.dll")]
        private static extern IntPtr ShowWindow(IntPtr hWnd, int nCmdShow); // 창의 상태를 설정

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowRect(IntPtr hWnd, ref Rect rect);  // 창의 크기와 위치를 가져옴

        /// <summary>
        /// 특정 프로세스의 창을 캡처하는 함수
        /// </summary>
        /// <param name="procName"></param>
        /// <returns></returns>
        public Bitmap CaptureApplication(string procName)
        {
            Process proc;   // 프로세스 객체 선언

            // 프로세스가 존재하지 않는 경우를 처리
            try
            {
                proc = Process.GetProcessesByName(procName)[0]; // 주어진 이름의 프로세스를 찾음, 현재는 Final_Project
            }
            catch (IndexOutOfRangeException e)
            {
                return null;    // 프로세스를 찾지 못하면 null 반환
            }

            // 창을 포커스하고 복원
            SetForegroundWindow(proc.MainWindowHandle);     // 창을 포커스로 설정
            ShowWindow(proc.MainWindowHandle, SW_RESTORE);  // 창을 복원 상태로 설정

            // 창이 복원되고 포커스되는 시간을 기다림
            Thread.Sleep(1000); // 1초 대기

            Rect rect = new Rect(); // 창의 위치와 크기를 저장할 구조체 생성
            IntPtr error = GetWindowRect(proc.MainWindowHandle, ref rect);  // 창의 위치와 크기 가져오기

            // GetWindowRect가 실패할 경우 반복 시도
            while (error == (IntPtr)0)
            {   // 성공할 때까지 반복
                error = GetWindowRect(proc.MainWindowHandle, ref rect);
            }

            int width = rect.right - rect.left;     // 창의 너비 계산
            int height = rect.bottom - rect.top;    // 창의 높이 계산

            // 비트맵 객체를 생성하고 스크린에서 이미지를 복사
            Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);    // 비트맵 객체 생성
            Graphics.FromImage(bmp).CopyFromScreen(rect.left, rect.top, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);  // 스크린에서 이미지를 복사하여 비트맵에 저장
            
            return bmp; // 캡처한 이미지를 반환
        }

        #endregion

        #region 이미지

        /// <summary>
        /// 이미지 경로를 리스트에 추가하는 함수
        /// </summary>
        private void LoadImages()
        {
            // 프로젝트 내부의 \bin\Debug\Images 폴더 경로 얻어오기
            // 제 컴퓨터 기준으로는 G:\4학년 1학기\윈도우 프로그래밍\과제\기말 프로젝트\Final_Project\bin\Debug\Images 입니다.
            string imagesFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images");

            // 폴더 경로에서 이미지를 얻어온 뒤 images List에 추가
            images.Add(Path.Combine(imagesFolder, "image01.png"));
            images.Add(Path.Combine(imagesFolder, "image02.png"));
            images.Add(Path.Combine(imagesFolder, "image03.png"));
        }

        /// <summary>
        /// 최초 로드 시 첫번째 이미지를 보여주는 함수
        /// </summary>
        private void ShowCurrentImage()
        {
            // iamges List에 이미지가 등록되어 있으면
            if (images.Count > 0)
            {
                // 홈 페이지의 pictureBox에 이미지 등록
                pictureBox2.Image = Image.FromFile(images[0]);
            }
        }

        #endregion

        #region 계정

        /// <summary>
        /// 계정 페이지에 등록하기 위해 모든 계정 정보를 읽어오는 함수
        /// </summary>
        async void GetALLDocumentsFromACollection()
        {
            // Firestore 데이터베이스의 "Add_Document_Width_AutoID" 컬렉션을 참조하는 쿼리 객체를 생성
            Query qref = db.Collection("Add_Document_Width_AutoID");
            QuerySnapshot snap = await qref.GetSnapshotAsync(); // 비동기적으로 쿼리 스냅샷을 가져옴

            // 가져온 스냅샷에서 각 문서를 반복하며
            foreach (DocumentSnapshot docsnap in snap)
            {
                // FirebaseProperty 객체로 변환
                FirebaseProperty fp = docsnap.ConvertTo<FirebaseProperty>();

                // 문서가 존재하고, 문서의 ID가 "Calendar", "Statistics"가 아닌 경우에만 리스트뷰에 추가
                if (docsnap.Exists && !docsnap.Id.Equals("Calendar") && !docsnap.Id.Equals("Statistics"))
                {
                    listView2.Items.Add(new ListViewItem(new string[] { docsnap.Id, fp.ID, fp.PW, fp.Su , fp.Name }));
                }
            }
        }

        /// <summary>
        /// 계정 생성 함수
        /// </summary>
        async void DataRegistration()
        {
            // 회원가입 폼의 Hashing 함수를 사용해서 고유 해시 값 생성
            string hash = form3.SHA256Hash($"{textBox10.Text.Trim()}{textBox11.Text.Trim()}");

            // Add_Document_Width_AutoID컬렉션에 해시 값 Document 생성
            DocumentReference coll = db.Collection("Add_Document_Width_AutoID").Document(hash);
            Dictionary<string, object> identityDocument = new Dictionary<string, object>()
            {
                { "ID", $"{textBox10.Text}" },
                { "PW", $"{textBox11.Text}" },
                { "Su", $"{domainUpDown1.Text}" },
                { "Name", $"{textBox12.Text}" },
            };  // Dictionary를 이용해 계정 정보 생성

            coll.SetAsync(identityDocument);    // 계정 정보 등록
        }

        #endregion

        #region 캘린더

        /// <summary>
        /// Firebase에서 Calendar를 읽어서 등록된 일정을 불러오거나 추가하는 함수
        /// </summary>
        async void GetMultipleDocumentsFromACollection()
        {
            // "Add_Document_Width_AutoID" 컬렉션의 "Calendar" 문서를 참조하는 DocumentReference 객체를 생성
            DocumentReference docref = db.Collection("Add_Document_Width_AutoID").Document("Calendar");
            DocumentSnapshot snap = await docref.GetSnapshotAsync();    // 비동기적으로 "Calendar" 문서의 스냅샷을 가져옴

            // monthCalendar1에서 선택된 날짜를 문자열 형식으로 가져옴
            string dateString = monthCalendar1.SelectionStart.ToShortDateString();

            // 문서가 존재하면
            if (snap.Exists)
            {
                // 문서를 Dictionary<string, object> 형식으로 변환
                Dictionary<string, object> schedule = snap.ToDictionary();

                object memo = string.Empty;

                // 선택된 날짜(dateString)가 schedule에 존재하는지 확인
                if (schedule.TryGetValue(dateString, out memo))
                {
                    // 존재하면 리스트박스에 날짜와 기존 메모에 새로운 메모를 추가하여 표시
                    listBox1.Items.Add($"{dateString} : {memo + "\n" + textBox13.Text}");
                    schedule[dateString] = memo + "\n" + textBox13.Text;
                }
                else
                {
                    // 존재하지 않으면 리스트박스에 날짜와 새로운 메모를 추가하여 표시
                    listBox1.Items.Add($"{dateString} : {textBox13.Text}");
                    schedule.Add(dateString, textBox13.Text);
                }

                // 업데이트된 schedule을 비동기적으로 Firestore에 저장
                docref.SetAsync(schedule);
                Clear(textBox13);
            }
        }

        /// <summary>
        /// Firebase에서 Calendar를 읽어 특정 날짜의 일정을 불러오는 함수
        /// </summary>
        async void GetCurrentDocumentsFromACollection()
        {
            // "Add_Document_Width_AutoID" 컬렉션의 "Calendar" 문서를 참조하는 DocumentReference 객체를 생성
            DocumentReference docref = db.Collection("Add_Document_Width_AutoID").Document("Calendar");
            DocumentSnapshot snap = await docref.GetSnapshotAsync();    // 비동기적으로 "Calendar" 문서의 스냅샷을 가져옴

            string dateString = monthCalendar1.SelectionStart.ToShortDateString();

            // 문서가 존재하면
            if (snap.Exists)
            {
                // 문서를 Dictionary<string, object> 형식으로 변환
                Dictionary<string, object> schedule = snap.ToDictionary();

                object memo = string.Empty;

                if (schedule.TryGetValue(dateString, out memo))
                {
                    listBox1.Items.Add($"{dateString} : {memo}");
                }
            }
        }

        /// <summary>
        /// Firebase의 Calendar에서 특정 날짜의 일정을 삭제하는 함수
        /// </summary>
        async void RemoveCurrentFieldFromCalendar()
        {
            // "Add_Document_Width_AutoID" 컬렉션의 "Calendar" 문서를 참조하는 DocumentReference 객체를 생성
            DocumentReference docref = db.Collection("Add_Document_Width_AutoID").Document("Calendar");
            DocumentSnapshot snap = await docref.GetSnapshotAsync();    // 비동기적으로 "Calendar" 문서의 스냅샷을 가져옴

            string dateString = monthCalendar1.SelectionStart.ToShortDateString();

            // 문서가 존재하면
            if (snap.Exists)
            {
                // 문서를 Dictionary<string, object> 형식으로 변환
                Dictionary<string, object> schedule = snap.ToDictionary();

                object memo = string.Empty;

                if (schedule.TryGetValue(dateString, out memo))
                {
                    schedule.Remove(dateString);
                }

                // 업데이트된 schedule을 비동기적으로 Firestore에 저장
                docref.SetAsync(schedule);
            }
        }

        /// <summary>
        /// Firebase의 Statistics에서 매출 값을 가져와 표시하는 함수
        /// </summary>
        async void GetStatisticsDocuments()
        {
            // "Add_Document_Width_AutoID" 컬렉션의 "Calendar" 문서를 참조하는 DocumentReference 객체를 생성
            DocumentReference docref = db.Collection("Add_Document_Width_AutoID").Document("Statistics");
            DocumentSnapshot snap = await docref.GetSnapshotAsync();    // 비동기적으로 "Statistics" 문서의 스냅샷을 가져옴

            string startDay = comboBox5.SelectedItem.ToString();
            string endDay = comboBox6.SelectedItem.ToString();

            // 문서가 존재하면
            if (snap.Exists)
            {
                // 문서를 Dictionary<string, object> 형식으로 변환
                Dictionary<string, object> statistics = snap.ToDictionary();

                Series sales = chart1.Series.FindByName("매출");

                sales.Points.Clear();

                object memo = string.Empty;

                // 시작, 종료 날짜 string을 DateTime으로 변환
                DateTime start = Convert.ToDateTime(startDay);
                DateTime end = Convert.ToDateTime(endDay);

                for (DateTime date = start; date <= end; date = date.AddDays(1))
                {
                    // 선택된 날짜(dateString)가 statistics에 존재하는지 확인
                    if (statistics.TryGetValue(date.ToShortDateString(), out memo))
                    {
                        sales.Points.AddXY(date.ToShortDateString(), int.Parse(memo.ToString()));
                    }
                    else
                    {
                        sales.Points.AddXY(date.ToShortDateString(), 0);
                    }
                }
            }
        }

        /// <summary>
        /// Firebase의 Statistics의 당일 매출에 결제 금액을 추가하는 함수
        /// </summary>
        async void AddStatistics()
        {
            // "Add_Document_Width_AutoID" 컬렉션의 "Calendar" 문서를 참조하는 DocumentReference 객체를 생성
            DocumentReference docref = db.Collection("Add_Document_Width_AutoID").Document("Statistics");
            DocumentSnapshot snap = await docref.GetSnapshotAsync();    // 비동기적으로 "Statistics" 문서의 스냅샷을 가져옴

            DateTime today = DateTime.Today;

            object memo = string.Empty;

            // 문서가 존재하면
            if (snap.Exists)
            {
                // 문서를 Dictionary<string, object> 형식으로 변환
                Dictionary<string, object> statistics = snap.ToDictionary();

                // 선택된 날짜(dateString)가 statistics에 존재하는지 확인
                if (statistics.TryGetValue(today.ToShortDateString(), out memo))
                {
                    statistics[today.ToShortDateString()] = int.Parse(memo.ToString()) + int.Parse(textBox8.Text);
                }
                else
                {
                    statistics.Add(today.ToShortDateString(), textBox8.Text);
                }

                // 업데이트된 schedule을 비동기적으로 Firestore에 저장
                docref.SetAsync(statistics);
                // 총 상품 금액 텍스트 박스 초기화
                Clear(textBox1, textBox2, textBox3, textBox4, textBox5, textBox6, textBox7, textBox8);
                listView1.Clear();
            }
        }

        #endregion

        #region 검색, 유튜브

        /// <summary>
        /// 폼 로드 시 웹 브라우저에 유튜브를 띄우는 함수
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            // HTML 문자열을 생성, IE 브라우저 호환성을 위해 메타 태그와 유튜브 영상을 포함하는 iframe을 포함
            var embed = "<html><head>" +
            "<meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge, chrome=1\"/>" +
            "</head><body>" +
            "<iframe width=\"440\" height=\"415\" src=\"{0}\"" +
            "frameborder = \"0\" allow = \"autoplay; encrypted-media\" allowfullscreen></iframe>" +
            "</body></html>";
            var url = "https://www.youtube.com/embed/kh52htRZsi8";      // 임의의 유튜브 링크 설정
            this.webBrowser1.DocumentText = string.Format(embed, url);  // 웹 브라우저 컨트롤에 HTML 문자열을 설정, 유튜브 URL을 embed 문자열의 {0} 위치에 삽입
        }

        /// <summary>
        /// 검색어를 크롬을 이용해 검색하는 함수
        /// </summary>
        /// <param name="url">검색할 쿼리가 추가된 최종 URL</param>
        private void OpenUrlInChrome(string url)
        {
            try
            {
                // 크롬을 이용해 전달받은 URL 사용
                Process.Start("chrome.exe", url);
            }
            catch (Exception ex)
            {
                // 크롬이 설치되지 않은 경우
                MessageBox.Show("크롬 브라우저를 열 수 없습니다. 크롬이 설치되어 있는지 확인해주세요.\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        /// <summary>
        /// 리스트 뷰에 상품을 추가하는 함수
        /// </summary>
        /// <param name="productInfo">QR코드에서 읽은 상품 정보</param>
        /// <returns>상품 정보를 기반으로 만든 ListView에 추가할 ListViewItem</returns>
        private ListViewItem ProductAdd(string productInfo)
        {
            string[] splitInfo = productInfo.Split(' ');    // 공백 기준으로 작성되어 있으므로 분리
            return new ListViewItem(new string[] { splitInfo[0], splitInfo[1], splitInfo[2], splitInfo[3], splitInfo[4], splitInfo[5], splitInfo[6] });
        }

        /// <summary>
        /// 입고 날짜 초기화 함수
        /// </summary>
        public void InitializeInStockDate()
        {
            // 현재 날짜로부터 일주일 이전까지의 입고날짜를 추가
            for (int i = 7; i >= 0; i--)
            {
                comboBox2.Items.Add(DateTime.Now.AddDays(-i).ToShortDateString());
                comboBox5.Items.Add(DateTime.Now.AddDays(-i).ToShortDateString());
            }

            // 현재 날짜로부터 일주일 이후까지의 입고날짜를 추가
            for (int i = 1; i < 7; i++)
            {
                comboBox2.Items.Add(DateTime.Now.AddDays(i).ToShortDateString());
                comboBox6.Items.Add(DateTime.Now.AddDays(i).ToShortDateString());
            }
        }

        /// <summary>
        /// 결제 페이지 총 결제 금액 최신화 함수
        /// </summary>
        public void UpdatePrice()
        {
            int totalPrice = 0;

            foreach (ListViewItem item in listView1.Items)
            {
                ListViewItem.ListViewSubItemCollection subItem = item.SubItems;
                totalPrice += int.Parse(subItem[4].Text) * int.Parse(subItem[5].Text);  // 단가 * 수량 
            }

            // 총 결제 금액 최신화
            textBox8.Text = totalPrice.ToString();
        }

        /// <summary>
        /// 프린트 함수
        /// </summary>
        /// <param name="listView">프린트 하려는 ListView</param>
        public void Print(ListView listView)
        {
            PageSettings pageSettings = new PageSettings();
            pageSetupDialog1.PageSettings = pageSettings;
            printDialog1.PrinterSettings = new PrinterSettings();   //프린터 설정
            printDialog1.Document = printDocument1; //인쇄 문서 설정

            DialogResult result = pageSetupDialog1.ShowDialog();    // 대화 상자 호출 및 사용자 입력 처리

            if (result == DialogResult.OK)
            {
                pageSettings = pageSetupDialog1.PageSettings;   // 변경된 설정 가져오기
                pageSet = $"\nMargins: {pageSettings.Margins}\nPaperSize: {pageSettings.PaperSize}";    // 설정된 정보 저장
                result = printDialog1.ShowDialog();

                if (result == DialogResult.OK)
                {
                    printDocument1.Print();
                }
            }
        }

        /// <summary>
        /// ListView를 받아서 엑셀로 옮기는 함수
        /// </summary>
        /// <param name="listView">엑셀로 옮길 ListView</param>
        public void SaveListViewToExcel(ListView listView = null)
        {
            // 엑셀 어플리케이션 생성
            Excel.Application excelApp = new Excel.Application();

            if (excelApp == null)
            {   // 엑셀 어플리케이션 생성 실패
                MessageBox.Show("Excel이 설치되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 엑셀 워크북과 워크시트 생성
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

            if (!(listView is null))
            {
                // 리스트 뷰의 헤더를 엑셀에 쓰기
                for (int i = 0; i < listView.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = listView.Columns[i].Text;
                }

                // 리스트 뷰의 아이템과 서브아이템을 엑셀에 쓰기
                for (int i = 0; i < listView.Items.Count; i++)
                {
                    for (int j = 0; j < listView.Items[i].SubItems.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = listView.Items[i].SubItems[j].Text;
                    }
                }
            }

            // 엑셀 보이기
            excelApp.Visible = true;
        }

        /// <summary>
        /// ListView 항목을 텍스트 박스로 출력하는 함수
        /// </summary>
        /// <param name="listView">텍스트 박스에 옮길 정보를 갖는 ListView</param>
        /// <param name="textBoxs">ListView정보를 출력할 텍스트 박스</param>
        public void ListViewToTextBox(ListView listView, params TextBox[] textBoxs)
        {   // ListView를 받고 TextBox를 가변 길이 매개변수로 받음
            foreach (ListViewItem item in listView.SelectedItems)
            {
                ListViewItem.ListViewSubItemCollection subItem = item.SubItems;

                for (int i = 0; i < subItem.Count; i++)
                {   // ListView의 값들을 매칭되는 TextBox에 넣어줌
                    textBoxs[i].Text = subItem[i].Text;
                }
            }
        }

        /// <summary>
        /// 입력 텍스트 박스가 비었는지 확인하는 함수
        /// </summary>
        /// <param name="textBoxs">비었는지 확인 할 텍스트 박스들</param>
        /// <returns>텍스트 박스가 하나라도 비었다면 True, 아니라면 False</returns>
        public bool IsEmpty(params TextBox[] textBoxs)
        {   // TextBox를 가변 길이 매개변수로 받고 텍스트가 비어있는지 확인
            foreach (TextBox textBox in textBoxs)
            {
                if (textBox.Text.Equals(string.Empty))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 입력 텍스트 박스 초기화 함수
        /// </summary>
        /// <param name="textBoxs">비우려는 텍스트 박스들</param>
        public void Clear(params TextBox[] textBoxs)
        {   // TextBox를 가변 길이 매개변수로 받고 모든 텍스트를 비움
            foreach (TextBox textBox in textBoxs)
            {
                textBox.Text = string.Empty;
            }
        }

        #endregion

        /***********************************************************************
        *                        좌측 탭 페이지 이동 버튼                        
        ***********************************************************************/
        #region .

        /// <summary>
        /// 홈 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage1;
        }

        /// <summary>
        /// 결제 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage2;
        }

        /// <summary>
        /// 재고 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage3;
        }

        /// <summary>
        /// 통계 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage4;
        }

        /// <summary>
        /// 계정 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage5;
        }

        /// <summary>
        /// 캡쳐 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button34_Click(object sender, EventArgs e)
        {
            // 캡쳐 함수에 현재 프로세스의 이름을 전달
            Bitmap bmpTmp = CaptureApplication("Final_Project");

            // 데스크탑 폴더 경로
            string deskPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            // 데스크탑 폴더 경로에 파일명 추가
            string filePath = Path.Combine(deskPath, $"Capture.jpg");

            // 캡쳐 파일 저장
            bmpTmp.Save(filePath, ImageFormat.Png);

            MessageBox.Show("캡쳐 사진이 저장되었습니다.", "캡쳐 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 종료 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e) => Close();

        #endregion

        /***********************************************************************
        *                             홈 페이지 로직                             
        ***********************************************************************/
        #region .

        /// <summary>
        /// 홈 페이지 시계 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            label14.Text = DateTime.Now.ToString("F");   // label14에 현재날짜 시간 표시, F : 자세한 전체 날짜/시간
        }

        #region 버튼

        /// <summary>
        /// 홈 페이지 계산기 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button24_Click(object sender, EventArgs e)
        {
            Process.Start("calc");
        }

        /// <summary>
        /// 홈 페이지 메모장 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button25_Click(object sender, EventArgs e)
        {
            Process.Start("notepad");
        }

        /// <summary>
        /// 홈 페이지 CMD 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button26_Click(object sender, EventArgs e)
        {
            Process.Start("cmd");
        }

        /// <summary>
        /// 홈 페이지 제어판 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button27_Click(object sender, EventArgs e)
        {
            Process.Start("control");
        }

        /// <summary>
        /// 홈 페이지 워드 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button22_Click(object sender, EventArgs e)
        {
            // 워드 어플리케이션 생성
            Word.Application wordApp = new Word.Application();

            if (wordApp == null)
            {
                // 워드 어플리케이션 생성 실패
                MessageBox.Show("Word가 설치되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 워드 문서 생성 후 보이게 설정
            wordApp.Documents.Add();
            wordApp.Visible = true;
        }

        /// <summary>
        /// 홈 페이지 엑셀 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button23_Click(object sender, EventArgs e)
        {
            SaveListViewToExcel();
        }

        /// <summary>
        /// 구글 검색 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button28_Click(object sender, EventArgs e)
        {
            string query = textBox16.Text;  // 텍스트 박스로부터 검색어를 얻어옴
            string url = "https://www.google.com/search?q=" + Uri.EscapeDataString(query);  // 구글 검색링크 생성
            OpenUrlInChrome(url);   // 실제 검색창 열기 함수로 URL 전달
        }

        #endregion

        #region 라디오 버튼

        /// <summary>
        /// 첫번째 이미지 보이기 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox2.Image = Image.FromFile(images[0]);
        }

        /// <summary>
        /// 두번째 이미지 보이기 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox2.Image = Image.FromFile(images[1]);
        }

        /// <summary>
        /// 세번째 이미지 보이기 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox2.Image = Image.FromFile(images[2]);
        }

        #endregion

        #region 캘린더

        /// <summary>
        /// 홈 페이지 캘린더 일정 등록 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button16_Click(object sender, EventArgs e)
        {
            // 출력전 남은 내용 삭제
            listBox1.Items.Clear();

            GetMultipleDocumentsFromACollection();
        }

        /// <summary>
        /// 홈 페이지 캘린더 일정 삭제 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button21_Click(object sender, EventArgs e)
        {
            // 현재 남은 내용 삭제
            listBox1.Items.Clear();

            RemoveCurrentFieldFromCalendar();
        }

        /// <summary>
        /// 캘린더 날짜 선택 시 호출 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            // 출력전 남은 내용 삭제
            listBox1.Items.Clear();

            GetCurrentDocumentsFromACollection();
        }

        #endregion

        #region 링크 라벨

        /// <summary>
        /// 개발자 깃허브 이동 링크 라벨 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://github.com/Lithium07z"); // 제 깃허브 링크입니다.
        }

        /// <summary>
        /// 한림대학교 링크 라벨 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://www.hallym.ac.kr/hallym_univ/");
        }

        /// <summary>
        /// 정보과학대학 링크 라벨 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://sw.hallym.ac.kr/");
        }

        /// <summary>
        /// SW사업단 링크 라벨 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://hlsw.hallym.ac.kr/index.php");
        }

        #endregion

        #endregion

        /***********************************************************************
        *                             결제 탭 페이지                             
        ***********************************************************************/
        #region .

        #region 버튼

        /// <summary>
        /// QR코드 인식 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            // ZXing 라이브러리를 사용하여 바코드를 읽는 BarcodeReader 객체를 생성
            ZXing.BarcodeReader barcodeReader = new ZXing.BarcodeReader();

            // 파일 경로를 저장할 문자열 초기화
            string filePath = string.Empty;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";       // 초기 디렉토리를 C 드라이브로 설정
            openFileDialog.Filter = "All Files | *.jpg";    // 파일 필터를 설정하여 .jpg 파일만 표시
            openFileDialog.CheckFileExists = true;  // 파일이 실제로 존재하는지 확인
            openFileDialog.CheckPathExists = true;  // 경로가 실제로 존재하는지 확인

            // 사용자가 파일을 선택하고 확인 버튼을 눌렀을 때
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // 선택한 파일의 경로를 filePath에 저장
                filePath = openFileDialog.FileName;

                // 선택한 파일로부터 비트맵 이미지를 생성
                Bitmap barcodeBitmap = (Bitmap)Image.FromFile(filePath);

                // 바코드 이미지를 디코딩하여 제품 정보를 추출
                var productInfo = barcodeReader.Decode(barcodeBitmap);

                // ListView에 제품 정보를 추가
                listView1.Items.Add(ProductAdd(productInfo.Text));
                
                // 가격 정보를 업데이트
                UpdatePrice();
            }
        }

        /// <summary>
        /// 상품 삭제 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.SelectedItems)
            {   // 선택된 모든 상품 삭제
                item.Remove();
            }

            UpdatePrice();  // 총 결제 금액 최신화
            Clear(textBox1, textBox2, textBox3, textBox4, textBox5, textBox6, textBox7);    // 입력창 지우기
        }

        /// <summary>
        /// 상품 수정 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {   // 선택된 모든 상품 가져오기
            foreach (ListViewItem item in listView1.SelectedItems)
            {   // 선택된 상품 정보를 현재 텍스트 박스의 값으로 수정
                ListViewItem.ListViewSubItemCollection subItem = item.SubItems;
                subItem[0].Text = textBox1.Text.Trim(); subItem[1].Text = textBox2.Text.Trim(); subItem[2].Text = textBox3.Text.Trim(); subItem[3].Text = textBox4.Text.Trim(); subItem[4].Text = textBox5.Text.Trim(); subItem[5].Text = textBox6.Text.Trim(); subItem[6].Text = textBox7.Text.Trim();
            }

            // 총 결제 금액 최신화
            UpdatePrice();
        }

        /// <summary>
        /// 상품 결제 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button10_Click(object sender, EventArgs e)
        {
            // 결제 금액을 매출에 추가
            AddStatistics();
        }

        /// <summary>
        /// 상품 등록 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button11_Click(object sender, EventArgs e)
        {
            if (IsEmpty(textBox1, textBox2, textBox3, textBox4, textBox5, textBox6, textBox7))
            {   // 텍스트 박스가 하나라도 비어있다면, 에러 발생
                MessageBox.Show("모든 옵션을 입력해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // ZXing 라이브러리를 사용하여 QR 코드를 생성하는 BarcodeWriter 객체를 생성
            ZXing.BarcodeWriter barcodeWriter = new ZXing.BarcodeWriter();
            barcodeWriter.Format = ZXing.BarcodeFormat.QR_CODE; // QR 코드 형식으로 설정

            // QR 코드 이미지의 너비와 높이를 PictureBox의 크기로 설정
            barcodeWriter.Options.Width = this.pictureBox1.Width;
            barcodeWriter.Options.Height = this.pictureBox1.Height;

            // 모든 텍스트 박스의 텍스트를 하나의 문자열로 결합하고, 앞뒤 공백을 제거
            string strQRCode = $"{textBox1.Text.Trim()} {textBox2.Text.Trim()} {textBox3.Text.Trim()} {textBox4.Text.Trim()} {textBox5.Text.Trim()} {textBox6.Text.Trim()} {textBox7.Text.Trim()}";

            // QR 코드를 생성하여 비트맵 이미지로 변환
            Bitmap QRCode = barcodeWriter.Write(strQRCode);

            // 생성된 QR 코드 이미지를 PictureBox에 설정
            this.pictureBox1.Image = QRCode;

            // 바탕화면 경로를 가져옴
            string deskPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            // QR 코드 이미지를 저장할 파일 경로를 설정 (텍스트 박스 3의 텍스트를 파일명으로 사용)
            string filePath = Path.Combine(deskPath, $"{textBox3.Text}.jpg");

            // QR 코드 이미지를 JPEG 형식으로 바탕화면에 저장
            QRCode.Save(filePath, ImageFormat.Jpeg);
        }

        #endregion

        /// <summary>
        /// 리스트뷰 항목 클릭 시 텍스트 뷰 출력 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_Click(object sender, EventArgs e)
        {
            ListViewToTextBox(listView1, textBox1, textBox2, textBox3, textBox4, textBox5, textBox6, textBox7);
        }

        #endregion

        /***********************************************************************
        *                            재고 페이지 로직                             
        ***********************************************************************/
        #region .

        #region 버튼

        /// <summary>
        /// 재고 등록 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button29_Click(object sender, EventArgs e)
        {
            if (IsEmpty(textBox17, textBox18, textBox19, textBox20, textBox21, textBox22, textBox23))
            {   // 텍스트 박스가 하나라도 비어있다면, 에러 발생
                MessageBox.Show("모든 옵션을 입력해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            listView3.Items.Add(ProductAdd($"{textBox17.Text} {textBox18.Text} {textBox19.Text} {textBox20.Text} {textBox21.Text} {textBox22.Text} {textBox23.Text}"));
            Clear(textBox17, textBox18, textBox19, textBox20, textBox21, textBox22, textBox23);
        }

        /// <summary>
        /// 재고 수정 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button30_Click(object sender, EventArgs e)
        {
            // 선택된 모든 상품 가져오기
            foreach (ListViewItem item in listView3.SelectedItems)
            {   // 선택된 상품 정보를 현재 텍스트 박스의 값으로 수정
                ListViewItem.ListViewSubItemCollection subItem = item.SubItems;
                subItem[0].Text = textBox17.Text.Trim(); subItem[1].Text = textBox18.Text.Trim(); subItem[2].Text = textBox19.Text.Trim(); subItem[3].Text = textBox20.Text.Trim(); subItem[4].Text = textBox21.Text.Trim(); subItem[5].Text = textBox22.Text.Trim(); subItem[6].Text = textBox23.Text.Trim();
            }
        }

        /// <summary>
        /// 재고 삭제 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button31_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView3.SelectedItems)
            {   // 선택된 모든 상품 삭제
                item.Remove();
            }

            Clear(textBox17, textBox18, textBox19, textBox20, textBox21, textBox22, textBox23); // 입력창 지우기
        }

        /// <summary>
        /// 엑셀 저장 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button32_Click(object sender, EventArgs e)
        {
            SaveListViewToExcel(listView3);
        }

        /// <summary>
        /// 인쇄 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button33_Click(object sender, EventArgs e)
        {
            listView = listView3;
            Print(listView3);
        }

        /// <summary>
        /// ListView에서 checkedListBox로 제품 이동 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button35_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView3.SelectedItems)
            {
                if (!checkedListBox1.Items.Contains(item.SubItems[2].Text))
                {
                    checkedListBox1.Items.Add(item.SubItems[2].Text);
                }
            }
        }

        /// <summary>
        /// checkedListBox에서 ListView로 제품 이동 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button36_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex != -1 && comboBox2.SelectedIndex != -1 && comboBox3.SelectedIndex != -1 && comboBox4.SelectedIndex != -1)
            {
                for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
                {
                    listView3.Items.Add(new ListViewItem(new string[] { checkedListBox1.CheckedIndices[i].ToString(), form3.SHA256Hash(checkedListBox1.CheckedItems[i].ToString()), checkedListBox1.CheckedItems[i].ToString(), comboBox3.SelectedItem.ToString(), comboBox1.SelectedItem.ToString(), comboBox4.SelectedItem.ToString(), comboBox2.SelectedItem.ToString() }));
                }

                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, false);
                }

                comboBox1.Text = comboBox2.Text = comboBox3.Text = comboBox4.Text = string.Empty;
            }
            else
            {
                MessageBox.Show("모든 옵션을 입력해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        /// <summary>
        /// 리스트뷰 항목 클릭 시 텍스트 뷰 출력 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView3_Click(object sender, EventArgs e)
        {
            ListViewToTextBox(listView3, textBox17, textBox18, textBox19, textBox20, textBox21, textBox22, textBox23);
        }

        #endregion

        /***********************************************************************
        *                            통계 페이지 로직                             
        ***********************************************************************/
        #region .

        /// <summary>
        /// 매출 통계 조회 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button19_Click(object sender, EventArgs e)
        {
            GetStatisticsDocuments();
        }

        #endregion

        /***********************************************************************
        *                            계정 페이지 로직                             
        ***********************************************************************/
        #region .

        /// <summary>
        /// 리스트뷰 항목 클릭 시 텍스트 뷰 출력 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView2_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView2.SelectedItems)
            {
                ListViewItem.ListViewSubItemCollection subItem = item.SubItems;
                textBox9.Text = subItem[0].Text;
                textBox10.Text = subItem[1].Text;
                textBox11.Text = subItem[2].Text;
                domainUpDown1.Text = subItem[3].Text;
                textBox12.Text = subItem[4].Text;
            }
        }

        #region 프린터

        /// <summary>
        /// 계정 정보 출력 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            Font printFont = new Font("Arial", 14, FontStyle.Bold);
            int index = 1;

            if (listView.Name.Equals("listView2"))
            {
                foreach (ListViewItem item in listView.Items)
                {
                    ListViewItem.ListViewSubItemCollection subItem = item.SubItems;
                    e.Graphics.DrawString("유저 코드 : " + subItem[0].Text + "\n아이디 : " + subItem[1].Text + "\n패스워드 : " + subItem[2].Text + "\n권한 : " + subItem[3].Text + "\n이름 : " + subItem[4].Text + "\n" + pageSet, printFont, Brushes.Black, 10, index * 10);
                    index += 20;
                }
            }
            else if (listView.Name.Equals("listView3"))
            {
                foreach (ListViewItem item in listView.Items)
                {
                    ListViewItem.ListViewSubItemCollection subItem = item.SubItems;
                    e.Graphics.DrawString("번호 : " + subItem[0].Text + "\n상품 코드 : " + subItem[1].Text + "\n상품명 : " + subItem[2].Text + "\n단가 : " + subItem[3].Text + "\n수량 : " + subItem[4].Text + "\n금액 : " + subItem[5].Text + "\n입고 : " + subItem[6].Text + "\n" + pageSet, printFont, Brushes.Black, 10, index * 10);
                    index += 20;
                }
            }
        }

        /// <summary>
        /// 프린트가 끝난 뒤 텍스트 박스 출력 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void printDocument1_EndPrint(object sender, PrintEventArgs e)
        {
            MessageBox.Show(printDocument1.DocumentName + " 인쇄 완료", "인쇄 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        #region 버튼

        /// <summary>
        /// 유저 등록 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button12_Click(object sender, EventArgs e)
        {   // 권한이 T인 경우만
            if (form2.su)
            {   
                listView2.Items.Add(new ListViewItem(new string[] { form3.SHA256Hash($"{textBox10.Text}{textBox11.Text}"), textBox10.Text, textBox11.Text, domainUpDown1.Text, textBox12.Text }));
                DataRegistration();
            }
        }

        /// <summary>
        /// 유저 수정 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button13_Click(object sender, EventArgs e)
        {
            if (form2.su)
            {
                if (listView2.SelectedItems.Count == 1)
                {
                    ListViewItem item = listView2.SelectedItems[0];
                    ListViewItem.ListViewSubItemCollection subItem = item.SubItems;
                    
                    DocumentReference coll = db.Collection("Add_Document_Width_AutoID").Document(subItem[0].Text);
                    coll.DeleteAsync();

                    string hash = form3.SHA256Hash($"{subItem[1].Text}{subItem[2].Text}");

                    coll = db.Collection("Add_Document_Width_AutoID").Document(hash);
                    Dictionary<string, object> identityDocument = new Dictionary<string, object>()
                    {
                        { "ID", $"{textBox10.Text}" },
                        { "PW", $"{textBox11.Text}" },
                        { "Su", $"{domainUpDown1.Text}" },
                        { "Name", $"{textBox12.Text}" },
                    };

                    coll.SetAsync(identityDocument);

                    textBox9.Text = subItem[0].Text = hash; subItem[1].Text = textBox10.Text.Trim(); subItem[2].Text = textBox11.Text.Trim(); subItem[3].Text = domainUpDown1.Text.Trim(); subItem[4].Text = textBox12.Text.Trim();
                }
            }
        }

        /// <summary>
        /// 유저 삭제 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button14_Click(object sender, EventArgs e)
        {
            if (form2.su && listView2.SelectedItems.Count > 0)
            {
                foreach (ListViewItem item in listView2.SelectedItems)
                {
                    ListViewItem.ListViewSubItemCollection subItem = item.SubItems;
                    DocumentReference coll = db.Collection("Add_Document_Width_AutoID").Document(subItem[0].Text);
                    coll.DeleteAsync();
                }

                listView2.SelectedItems[0].Remove();

                Clear(textBox9, textBox10, textBox11, textBox12);
                domainUpDown1.Text = "T";
            }
        }

        /// <summary>
        /// 엑셀 저장 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button15_Click(object sender, EventArgs e)
        {
            SaveListViewToExcel(listView2);
        }

        /// <summary>
        /// 인쇄 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button17_Click(object sender, EventArgs e)
        {
            listView = listView2;
            Print(listView2);
        }

        /// <summary>
        /// 초기화 버튼 함수
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button18_Click(object sender, EventArgs e)
        {
            Clear(textBox9, textBox10, textBox11, textBox12);
            domainUpDown1.Text = "T";
        }

        #endregion

        #endregion

    }
}
