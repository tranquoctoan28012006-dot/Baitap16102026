using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ClinicAppointmentManager;

public class Appointment
{
    public string MaLichHen { get; set; } = "";
    public string HoTen { get; set; } = "";
    public string SoDienThoai { get; set; } = "";
    public DateTime NgayGioHen { get; set; }
    public string BacSi { get; set; } = "";
    public string LyDoKham { get; set; } = "";
    public string TrangThai { get; set; } = "Chờ khám";
}

public class MainForm : Form
{
    private readonly BindingList<Appointment> items=new();
    private readonly BindingSource source=new();
    private readonly BindingSource viewSource=new();
    private readonly ErrorProvider errors=new();
    private readonly DataGridView grid=new();
    private readonly TextBox txtSearch=new();
    private readonly ComboBox filterStatus=new();
    private readonly DateTimePicker filterDate=new();
    private readonly ToolStripStatusLabel lblStatus=new(){Text="Sẵn sàng"};
    private readonly TableLayoutPanel fields=new();
    private readonly TextBox txtId=new(),txtName=new(),txtPhone=new(),txtReason=new(); private readonly ComboBox cboDoctor=new(),cboStatus=new(); private readonly DateTimePicker dtMain=new();
    private Appointment? selected;

    public MainForm()
    {
        Text="ClinicAppointmentManager"; Width=1250; Height=780; MinimumSize=new Size(900,600); StartPosition=FormStartPosition.CenterScreen;
        BuildInterface();
        source.DataSource=items; grid.DataSource=viewSource;
        items.Add(new Appointment{MaLichHen="LH001",HoTen="Nguyễn Minh Anh",SoDienThoai="0912345678",NgayGioHen=DateTime.Now.AddHours(2),BacSi="BS. Nguyễn An",LyDoKham="Khám tổng quát",TrangThai="Chờ khám"});
        RefreshView();
    }
    private void BuildInterface()
    {
        var menu=new MenuStrip();var file=new ToolStripMenuItem("Tệp");var exp=new ToolStripMenuItem("Xuất CSV");
        exp.Click+=(_,__)=>ExportCsv();file.DropDownItems.Add(exp);menu.Items.Add(file);MainMenuStrip=menu;Controls.Add(menu);
        var status=new StatusStrip();status.Items.Add(lblStatus);Controls.Add(status);
        var root=new TableLayoutPanel{Dock=DockStyle.Fill,Padding=new Padding(8),ColumnCount=1,RowCount=2};
        root.RowStyles.Add(new RowStyle(SizeType.Absolute,300));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));Controls.Add(root);root.BringToFront();
        var inputPanel=new Panel{Dock=DockStyle.Fill,AutoScroll=true};
        fields.Dock=DockStyle.Top;fields.AutoSize=true;fields.ColumnCount=2;fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,150));fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));inputPanel.Controls.Add(fields);
        AddText("Mã lịch hẹn",txtId); AddText("Họ tên bệnh nhân",txtName); AddText("Số điện thoại",txtPhone); AddCombo("Bác sĩ",cboDoctor,new[]{"BS. Nguyễn An","BS. Trần Bình","BS. Lê Chi","BS. Phạm Dũng"}); AddText("Lý do khám",txtReason); AddDate("Ngày giờ hẹn",dtMain,true); AddCombo("Trạng thái",cboStatus,new[]{"Chờ khám","Đang khám","Hoàn tất","Đã hủy"});
        
        var buttons=new FlowLayoutPanel{Dock=DockStyle.Bottom,Height=42};
        buttons.Controls.Add(MakeButton("Thêm",(_,__)=>AddItem()));buttons.Controls.Add(MakeButton("Cập nhật",(_,__)=>UpdateItem()));buttons.Controls.Add(MakeButton("Xóa",(_,__)=>DeleteItem()));buttons.Controls.Add(MakeButton("Làm mới",(_,__)=>ClearInputs()));
        inputPanel.Controls.Add(buttons);root.Controls.Add(inputPanel,0,0);
        var list=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2};list.RowStyles.Add(new RowStyle(SizeType.Absolute,42));list.RowStyles.Add(new RowStyle(SizeType.Percent,100));
        var searchPanel=new FlowLayoutPanel{Dock=DockStyle.Fill,WrapContents=false,AutoScroll=true};
        searchPanel.Controls.Add(new Label{Text="Tìm kiếm:",AutoSize=true,Padding=new Padding(0,8,0,0)});txtSearch.Width=220;txtSearch.TextChanged+=(_,__)=>RefreshView();searchPanel.Controls.Add(txtSearch);
        searchPanel.Controls.Add(new Label{Text="Lọc:",AutoSize=true,Padding=new Padding(8,8,0,0)});filterStatus.DropDownStyle=ComboBoxStyle.DropDownList;filterStatus.Width=170;filterStatus.Items.AddRange(new[]{"Tất cả","Chờ khám","Đang khám","Hoàn tất","Đã hủy"});filterStatus.SelectedIndex=0;filterStatus.SelectedIndexChanged+=(_,__)=>RefreshView();searchPanel.Controls.Add(filterStatus);
        filterDate.ShowCheckBox=true;filterDate.Checked=false;filterDate.Format=DateTimePickerFormat.Short;searchPanel.Controls.Add(filterDate);filterDate.ValueChanged+=(_,__)=>RefreshView();
        list.Controls.Add(searchPanel,0,0);
        grid.Dock=DockStyle.Fill;grid.ReadOnly=true;grid.AllowUserToAddRows=false;grid.AutoGenerateColumns=false;grid.SelectionMode=DataGridViewSelectionMode.FullRowSelect;grid.MultiSelect=false;grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill;grid.RowHeadersVisible=false;grid.SelectionChanged+=(_,__)=>LoadSelected();
        AddColumn("Mã lịch hẹn","MaLichHen");
        AddColumn("Họ tên","HoTen");
        AddColumn("Số điện thoại","SoDienThoai");
        AddColumn("Ngày giờ hẹn","NgayGioHen", "dd/MM/yyyy HH:mm");
        AddColumn("Bác sĩ","BacSi");
        AddColumn("Lý do khám","LyDoKham");
        AddColumn("Trạng thái","TrangThai");
        
        list.Controls.Add(grid,0,1);root.Controls.Add(list,0,1);errors.ContainerControl=this;
    }
    private void AddText(string label,TextBox c){c.Dock=DockStyle.Fill;AddField(label,c);}
    private void AddCombo(string label,ComboBox c,string[] values){c.DropDownStyle=ComboBoxStyle.DropDownList;c.Dock=DockStyle.Fill;c.Items.AddRange(values);if(c.Items.Count>0)c.SelectedIndex=0;AddField(label,c);}
    private void AddDate(string label,DateTimePicker c,bool time){c.Dock=DockStyle.Fill;c.Format=time?DateTimePickerFormat.Custom:DateTimePickerFormat.Short;if(time)c.CustomFormat="dd/MM/yyyy HH:mm";AddField(label,c);}
    private void AddNumber(string label,NumericUpDown c,decimal min,decimal max){c.Minimum=min;c.Maximum=max;c.Dock=DockStyle.Fill;AddField(label,c);}
    private void AddField(string label,Control c){int r=fields.RowCount++;fields.RowStyles.Add(new RowStyle(SizeType.Absolute,36));fields.Controls.Add(new Label{Text=label,TextAlign=ContentAlignment.MiddleLeft,Dock=DockStyle.Fill},0,r);fields.Controls.Add(c,1,r);}
    private void AddColumn(string title,string prop,string? fmt=null){var c=new DataGridViewTextBoxColumn{HeaderText=title,DataPropertyName=prop,MinimumWidth=75};if(fmt!=null)c.DefaultCellStyle.Format=fmt;grid.Columns.Add(c);}
    private Button MakeButton(string text,EventHandler handler){var b=new Button{Text=text,Width=100,Height=30,Margin=new Padding(4)};b.Click+=handler;return b;}
    private bool ValidateInput()
    {
        errors.Clear();bool ok=true;
        if(string.IsNullOrWhiteSpace(txtId.Text)){errors.SetError(txtId,"Bắt buộc nhập mã.");ok=false;}
        if(items.Any(x=>GetId(x).Equals(txtId.Text.Trim(),StringComparison.OrdinalIgnoreCase)&&x!=selected)){errors.SetError(txtId,"Mã đã tồn tại.");ok=false;}
        if(string.IsNullOrWhiteSpace(txtName.Text)){errors.SetError(txtName,"Bắt buộc nhập họ tên.");ok=false;} if(!System.Text.RegularExpressions.Regex.IsMatch(txtPhone.Text.Trim(),@"^(0\d{9,10}|\+84\d{9,10})$")){errors.SetError(txtPhone,"Số điện thoại không hợp lệ.");ok=false;} if(string.IsNullOrWhiteSpace(txtReason.Text)){errors.SetError(txtReason,"Bắt buộc nhập lý do khám.");ok=false;} if(dtMain.Value<DateTime.Now.AddMinutes(-1)){errors.SetError(dtMain,"Ngày hẹn không được trong quá khứ.");ok=false;}
        return ok;
    }
    private string GetId(Appointment x) => x.MaLichHen;
    private void AddItem(){if(!ValidateInput())return;items.Add(new Appointment{MaLichHen=txtId.Text.Trim(),HoTen=txtName.Text.Trim(),SoDienThoai=txtPhone.Text.Trim(),NgayGioHen=dtMain.Value,BacSi=cboDoctor.Text,LyDoKham=txtReason.Text.Trim(),TrangThai=cboStatus.Text});RefreshView();ClearInputs();}
    private void UpdateItem(){if(selected==null){MessageBox.Show("Hãy chọn bản ghi cần cập nhật.");return;}if(!ValidateInput())return;var a=selected;a.MaLichHen=txtId.Text.Trim();a.HoTen=txtName.Text.Trim();a.SoDienThoai=txtPhone.Text.Trim();a.NgayGioHen=dtMain.Value;a.BacSi=cboDoctor.Text;a.LyDoKham=txtReason.Text.Trim();a.TrangThai=cboStatus.Text;source.ResetBindings(false);RefreshView();}
    private void DeleteItem(){if(selected==null){MessageBox.Show("Hãy chọn bản ghi cần xóa.");return;}if(MessageBox.Show("Bạn có chắc muốn xóa?","Xác nhận",MessageBoxButtons.YesNo,MessageBoxIcon.Warning)!=DialogResult.Yes)return;items.Remove(selected);selected=null;ClearInputs();RefreshView();}
    private void ClearInputs(){selected=null;txtId.Clear();txtName.Clear();txtPhone.Clear();txtReason.Clear();cboDoctor.SelectedIndex=0;cboStatus.SelectedIndex=0;dtMain.Value=DateTime.Now.AddHours(1);errors.Clear();grid.ClearSelection();}
    private void LoadSelected(){if(grid.CurrentRow?.DataBoundItem is not Appointment a)return;selected=a;txtId.Text=a.MaLichHen;txtName.Text=a.HoTen;txtPhone.Text=a.SoDienThoai;txtReason.Text=a.LyDoKham;cboDoctor.Text=a.BacSi;cboStatus.Text=a.TrangThai;dtMain.Value=a.NgayGioHen;}
    private void RefreshView(){var q=txtSearch.Text.Trim();var view=items.Where(x=>(x.MaLichHen.Contains(q,StringComparison.OrdinalIgnoreCase)||x.HoTen.Contains(q,StringComparison.OrdinalIgnoreCase)||x.SoDienThoai.Contains(q,StringComparison.OrdinalIgnoreCase))&&((filterStatus.SelectedIndex==0||x.TrangThai==filterStatus.Text)&&(!filterDate.Checked||x.NgayGioHen.Date==filterDate.Value.Date))).ToList();viewSource.DataSource=view;lblStatus.Text=$"Số lịch hẹn sau lọc: {view.Count}";}
    private void ExportCsv(){using var d=new SaveFileDialog{Filter="CSV (*.csv)|*.csv",FileName="ClinicAppointmentManager.csv"};if(d.ShowDialog()!=DialogResult.OK)return;var q=txtSearch.Text.Trim();var view=items.Where(x=>(x.MaLichHen.Contains(q,StringComparison.OrdinalIgnoreCase)||x.HoTen.Contains(q,StringComparison.OrdinalIgnoreCase)||x.SoDienThoai.Contains(q,StringComparison.OrdinalIgnoreCase))&&((filterStatus.SelectedIndex==0||x.TrangThai==filterStatus.Text)&&(!filterDate.Checked||x.NgayGioHen.Date==filterDate.Value.Date))).ToList();using var w=new StreamWriter(d.FileName,false,new UTF8Encoding(true));w.WriteLine(string.Join(",",new string[]{"Mã lịch hẹn","Họ tên","Số điện thoại","Ngày giờ hẹn","Bác sĩ","Lý do khám","Trạng thái"}.Select(Csv)));foreach(var a in view)w.WriteLine(string.Join(",",new[]{a.MaLichHen,a.HoTen,a.SoDienThoai,a.NgayGioHen.ToString("yyyy-MM-dd HH:mm"),a.BacSi,a.LyDoKham,a.TrangThai}.Select(Csv)));MessageBox.Show("Xuất CSV thành công.");}
    private static string Csv(string s) => "\"" + (s ?? "").Replace("\"", "\"\"") + "\"";
    
}
