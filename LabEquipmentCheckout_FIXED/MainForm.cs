using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LabEquipmentCheckout;

public class Checkout
{
    public string MaPhieu { get; set; } = "";
    public string MaSinhVien { get; set; } = "";
    public string HoTenSinhVien { get; set; } = "";
    public string ThietBi { get; set; } = "";
    public int SoLuong { get; set; }
    public DateTime NgayMuon { get; set; }
    public DateTime HanTra { get; set; }
    public DateTime? NgayTraThucTe { get; set; }
    public string TinhTrang { get; set; } = "Đang mượn";
}

public class MainForm : Form
{
    private readonly BindingList<Checkout> items=new();
    private readonly BindingSource source=new();
    private readonly BindingSource viewSource=new();
    private readonly ErrorProvider errors=new();
    private readonly DataGridView grid=new();
    private readonly TextBox txtSearch=new();
    private readonly ComboBox filterStatus=new();
    private readonly DateTimePicker filterDate=new();
    private readonly ToolStripStatusLabel lblStatus=new(){Text="Sẵn sàng"};
    private readonly TableLayoutPanel fields=new();
    private readonly TextBox txtId=new(),txtStudentId=new(),txtName=new(); private readonly ComboBox cboDevice=new(),cboStatus=new(); private readonly NumericUpDown numQty=new(){Minimum=1,Maximum=999,Value=1}; private readonly DateTimePicker dtBorrow=new(),dtDue=new(),dtReturned=new();
    private Checkout? selected;

    public MainForm()
    {
        Text="LabEquipmentCheckout"; Width=1250; Height=780; MinimumSize=new Size(900,600); StartPosition=FormStartPosition.CenterScreen;
        BuildInterface();
        source.DataSource=items; grid.DataSource=viewSource;
        items.Add(new Checkout{MaPhieu="MP001",MaSinhVien="SV001",HoTenSinhVien="Trần Minh",ThietBi="Laptop",SoLuong=1,NgayMuon=DateTime.Today.AddDays(-2),HanTra=DateTime.Today.AddDays(5),TinhTrang="Đang mượn"});
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
        AddText("Mã phiếu",txtId); AddText("Mã sinh viên",txtStudentId); AddText("Họ tên sinh viên",txtName); AddCombo("Thiết bị",cboDevice,new[]{"Laptop","Máy chiếu","Arduino","Đồng hồ vạn năng","Bộ thí nghiệm"}); AddNumber("Số lượng",numQty,1,999); AddDate("Ngày mượn",dtBorrow,false); AddDate("Hạn trả",dtDue,false); AddDate("Ngày trả thực tế",dtReturned,false); AddCombo("Tình trạng",cboStatus,new[]{"Đang mượn","Đã trả","Quá hạn"});
        
        var buttons=new FlowLayoutPanel{Dock=DockStyle.Bottom,Height=42};
        buttons.Controls.Add(MakeButton("Thêm",(_,__)=>AddItem()));buttons.Controls.Add(MakeButton("Cập nhật",(_,__)=>UpdateItem()));buttons.Controls.Add(MakeButton("Xóa",(_,__)=>DeleteItem()));buttons.Controls.Add(MakeButton("Làm mới",(_,__)=>ClearInputs()));
        inputPanel.Controls.Add(buttons);root.Controls.Add(inputPanel,0,0);
        var list=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2};list.RowStyles.Add(new RowStyle(SizeType.Absolute,42));list.RowStyles.Add(new RowStyle(SizeType.Percent,100));
        var searchPanel=new FlowLayoutPanel{Dock=DockStyle.Fill,WrapContents=false,AutoScroll=true};
        searchPanel.Controls.Add(new Label{Text="Tìm kiếm:",AutoSize=true,Padding=new Padding(0,8,0,0)});txtSearch.Width=220;txtSearch.TextChanged+=(_,__)=>RefreshView();searchPanel.Controls.Add(txtSearch);
        searchPanel.Controls.Add(new Label{Text="Lọc:",AutoSize=true,Padding=new Padding(8,8,0,0)});filterStatus.DropDownStyle=ComboBoxStyle.DropDownList;filterStatus.Width=170;filterStatus.Items.AddRange(new[]{"Tất cả","Đang mượn","Quá hạn","Đã trả"});filterStatus.SelectedIndex=0;filterStatus.SelectedIndexChanged+=(_,__)=>RefreshView();searchPanel.Controls.Add(filterStatus);
        
        list.Controls.Add(searchPanel,0,0);
        grid.Dock=DockStyle.Fill;grid.ReadOnly=true;grid.AllowUserToAddRows=false;grid.AutoGenerateColumns=false;grid.SelectionMode=DataGridViewSelectionMode.FullRowSelect;grid.MultiSelect=false;grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill;grid.RowHeadersVisible=false;grid.SelectionChanged+=(_,__)=>LoadSelected();
        AddColumn("Mã phiếu","MaPhieu");
        AddColumn("Mã sinh viên","MaSinhVien");
        AddColumn("Họ tên sinh viên","HoTenSinhVien");
        AddColumn("Thiết bị","ThietBi");
        AddColumn("Số lượng","SoLuong");
        AddColumn("Ngày mượn","NgayMuon", "dd/MM/yyyy");
        AddColumn("Hạn trả","HanTra", "dd/MM/yyyy");
        AddColumn("Ngày trả thực tế","NgayTraThucTe", "dd/MM/yyyy");
        AddColumn("Tình trạng","TinhTrang");
        
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
        if(string.IsNullOrWhiteSpace(txtStudentId.Text)){errors.SetError(txtStudentId,"Bắt buộc mã sinh viên.");ok=false;}if(string.IsNullOrWhiteSpace(cboDevice.Text)){errors.SetError(cboDevice,"Chọn thiết bị.");ok=false;}if(dtDue.Value.Date<dtBorrow.Value.Date){errors.SetError(dtDue,"Hạn trả không được trước ngày mượn.");ok=false;}if(cboStatus.Text=="Đã trả"&&dtReturned.Value.Date<dtBorrow.Value.Date){errors.SetError(dtReturned,"Ngày trả không được trước ngày mượn.");ok=false;}
        return ok;
    }
    private string GetId(Checkout x) => x.MaPhieu;
    private void AddItem(){if(!ValidateInput())return;items.Add(new Checkout{MaPhieu=txtId.Text.Trim(),MaSinhVien=txtStudentId.Text.Trim(),HoTenSinhVien=txtName.Text.Trim(),ThietBi=cboDevice.Text,SoLuong=(int)numQty.Value,NgayMuon=dtBorrow.Value.Date,HanTra=dtDue.Value.Date,NgayTraThucTe=cboStatus.Text=="Đã trả"?dtReturned.Value.Date:null,TinhTrang=cboStatus.Text});RefreshView();ClearInputs();}
    private void UpdateItem(){if(selected==null){MessageBox.Show("Hãy chọn bản ghi cần cập nhật.");return;}if(!ValidateInput())return;var a=selected;a.MaPhieu=txtId.Text.Trim();a.MaSinhVien=txtStudentId.Text.Trim();a.HoTenSinhVien=txtName.Text.Trim();a.ThietBi=cboDevice.Text;a.SoLuong=(int)numQty.Value;a.NgayMuon=dtBorrow.Value.Date;a.HanTra=dtDue.Value.Date;a.NgayTraThucTe=cboStatus.Text=="Đã trả"?dtReturned.Value.Date:null;a.TinhTrang=cboStatus.Text;source.ResetBindings(false);RefreshView();}
    private void DeleteItem(){if(selected==null){MessageBox.Show("Hãy chọn bản ghi cần xóa.");return;}if(MessageBox.Show("Bạn có chắc muốn xóa?","Xác nhận",MessageBoxButtons.YesNo,MessageBoxIcon.Warning)!=DialogResult.Yes)return;items.Remove(selected);selected=null;ClearInputs();RefreshView();}
    private void ClearInputs(){selected=null;txtId.Clear();txtStudentId.Clear();txtName.Clear();cboDevice.SelectedIndex=0;numQty.Value=1;dtBorrow.Value=DateTime.Today;dtDue.Value=DateTime.Today.AddDays(7);dtReturned.Value=DateTime.Today;cboStatus.SelectedIndex=0;errors.Clear();grid.ClearSelection();}
    private void LoadSelected(){if(grid.CurrentRow?.DataBoundItem is not Checkout a)return;selected=a;txtId.Text=a.MaPhieu;txtStudentId.Text=a.MaSinhVien;txtName.Text=a.HoTenSinhVien;cboDevice.Text=a.ThietBi;numQty.Value=Math.Max(1,a.SoLuong);dtBorrow.Value=a.NgayMuon;dtDue.Value=a.HanTra;dtReturned.Value=a.NgayTraThucTe??DateTime.Today;cboStatus.Text=a.TinhTrang;}
    private void RefreshView(){var q=txtSearch.Text.Trim();var view=items.Where(x=>(x.MaSinhVien.Contains(q,StringComparison.OrdinalIgnoreCase)||x.ThietBi.Contains(q,StringComparison.OrdinalIgnoreCase)||x.HoTenSinhVien.Contains(q,StringComparison.OrdinalIgnoreCase)||x.MaPhieu.Contains(q,StringComparison.OrdinalIgnoreCase))&&((filterStatus.SelectedIndex==0||x.TinhTrang==filterStatus.Text))).ToList();viewSource.DataSource=view;lblStatus.Text=$"Tổng phiếu: {view.Count} | Quá hạn: {view.Count(x=>x.TinhTrang=="Quá hạn")}";}
    private void ExportCsv(){using var d=new SaveFileDialog{Filter="CSV (*.csv)|*.csv",FileName="LabEquipmentCheckout.csv"};if(d.ShowDialog()!=DialogResult.OK)return;var q=txtSearch.Text.Trim();var view=items.Where(x=>(x.MaSinhVien.Contains(q,StringComparison.OrdinalIgnoreCase)||x.ThietBi.Contains(q,StringComparison.OrdinalIgnoreCase)||x.HoTenSinhVien.Contains(q,StringComparison.OrdinalIgnoreCase)||x.MaPhieu.Contains(q,StringComparison.OrdinalIgnoreCase))&&((filterStatus.SelectedIndex==0||x.TinhTrang==filterStatus.Text))).ToList();using var w=new StreamWriter(d.FileName,false,new UTF8Encoding(true));w.WriteLine(string.Join(",",new string[]{"Mã phiếu","Mã sinh viên","Họ tên sinh viên","Thiết bị","Số lượng","Ngày mượn","Hạn trả","Ngày trả thực tế","Tình trạng"}.Select(Csv)));foreach(var a in view)w.WriteLine(string.Join(",",new[]{a.MaPhieu,a.MaSinhVien,a.HoTenSinhVien,a.ThietBi,a.SoLuong.ToString(),a.NgayMuon.ToString("yyyy-MM-dd"),a.HanTra.ToString("yyyy-MM-dd"),a.NgayTraThucTe?.ToString("yyyy-MM-dd")??"",a.TinhTrang}.Select(Csv)));MessageBox.Show("Xuất CSV thành công.");}
    private static string Csv(string s) => "\"" + (s ?? "").Replace("\"", "\"\"") + "\"";
    
}
