using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RepairServiceTracker;

public class RepairOrder
{
    public string MaDon { get; set; } = "";
    public string TenKhachHang { get; set; } = "";
    public string SoDienThoai { get; set; } = "";
    public string LoaiThietBi { get; set; } = "";
    public string MoTaLoi { get; set; } = "";
    public DateTime NgayTiepNhan { get; set; }
    public decimal ChiPhiDuKien { get; set; }
    public string TrangThai { get; set; } = "Mới tiếp nhận";
    public string DuongDanAnh { get; set; } = "";
}

public class MainForm : Form
{
    private readonly BindingList<RepairOrder> items=new();
    private readonly BindingSource source=new();
    private readonly BindingSource viewSource=new();
    private readonly ErrorProvider errors=new();
    private readonly DataGridView grid=new();
    private readonly TextBox txtSearch=new();
    private readonly ComboBox filterStatus=new();
    private readonly DateTimePicker filterDate=new();
    private readonly ToolStripStatusLabel lblStatus=new(){Text="Sẵn sàng"};
    private readonly TableLayoutPanel fields=new();
    private readonly TextBox txtId=new(),txtName=new(),txtPhone=new(),txtDescription=new(); private readonly ComboBox cboDevice=new(),cboStatus=new(); private readonly DateTimePicker dtMain=new(); private readonly NumericUpDown numCost=new(){Minimum=0,Maximum=1000000000,DecimalPlaces=0,ThousandsSeparator=true}; private readonly PictureBox pic=new(){Width=130,Height=90,SizeMode=PictureBoxSizeMode.Zoom,BorderStyle=BorderStyle.FixedSingle}; private string imagePath="";
    private RepairOrder? selected;

    public MainForm()
    {
        Text="RepairServiceTracker"; Width=1250; Height=780; MinimumSize=new Size(900,600); StartPosition=FormStartPosition.CenterScreen;
        BuildInterface();
        source.DataSource=items; grid.DataSource=viewSource;
        items.Add(new RepairOrder{MaDon="SC001",TenKhachHang="Nguyễn Hà",SoDienThoai="0987654321",LoaiThietBi="Laptop",MoTaLoi="Không lên nguồn",NgayTiepNhan=DateTime.Today,ChiPhiDuKien=850000,TrangThai="Mới tiếp nhận"});
        RefreshView();
    }
    private void BuildInterface()
    {
        var menu=new MenuStrip();var file=new ToolStripMenuItem("Tệp");var exp=new ToolStripMenuItem("Xuất CSV");
        exp.Click+=(_,__)=>ExportCsv();file.DropDownItems.Add(exp);file.DropDownItems.Add(new ToolStripMenuItem("Thoát",null,(_,__)=>Close()));menu.Items.Add(file);MainMenuStrip=menu;Controls.Add(menu);
        var status=new StatusStrip();status.Items.Add(lblStatus);Controls.Add(status);
        var root=new TableLayoutPanel{Dock=DockStyle.Fill,Padding=new Padding(8),ColumnCount=1,RowCount=2};
        root.RowStyles.Add(new RowStyle(SizeType.Absolute,300));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));Controls.Add(root);root.BringToFront();
        var inputPanel=new Panel{Dock=DockStyle.Fill,AutoScroll=true};
        fields.Dock=DockStyle.Top;fields.AutoSize=true;fields.ColumnCount=2;fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,150));fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));inputPanel.Controls.Add(fields);
        AddText("Mã đơn",txtId); AddText("Tên khách hàng",txtName); AddText("Số điện thoại",txtPhone); AddCombo("Loại thiết bị",cboDevice,new[]{"Điện thoại","Laptop","Máy tính bảng","Máy tính bàn","Thiết bị khác"}); AddText("Mô tả lỗi",txtDescription); AddDate("Ngày tiếp nhận",dtMain,false); AddNumber("Chi phí dự kiến",numCost,0,1000000000); AddCombo("Trạng thái",cboStatus,new[]{"Mới tiếp nhận","Đang sửa","Chờ linh kiện","Hoàn tất","Đã hủy"});
        inputPanel.Controls.Add(pic);
        var buttons=new FlowLayoutPanel{Dock=DockStyle.Bottom,Height=42};
        buttons.Controls.Add(MakeButton("Thêm",(_,__)=>AddItem()));buttons.Controls.Add(MakeButton("Cập nhật",(_,__)=>UpdateItem()));buttons.Controls.Add(MakeButton("Xóa",(_,__)=>DeleteItem()));buttons.Controls.Add(MakeButton("Làm mới",(_,__)=>ClearInputs()));var btnImage=MakeButton("Chọn ảnh",(_,__)=>ChooseImage()); buttons.Controls.Add(btnImage);
        inputPanel.Controls.Add(buttons);root.Controls.Add(inputPanel,0,0);
        var list=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2};list.RowStyles.Add(new RowStyle(SizeType.Absolute,42));list.RowStyles.Add(new RowStyle(SizeType.Percent,100));
        var searchPanel=new FlowLayoutPanel{Dock=DockStyle.Fill,WrapContents=false,AutoScroll=true};
        searchPanel.Controls.Add(new Label{Text="Tìm kiếm:",AutoSize=true,Padding=new Padding(0,8,0,0)});txtSearch.Width=220;txtSearch.TextChanged+=(_,__)=>RefreshView();searchPanel.Controls.Add(txtSearch);
        searchPanel.Controls.Add(new Label{Text="Lọc:",AutoSize=true,Padding=new Padding(8,8,0,0)});filterStatus.DropDownStyle=ComboBoxStyle.DropDownList;filterStatus.Width=170;filterStatus.Items.AddRange(new[]{"Tất cả","Mới tiếp nhận","Đang sửa","Chờ linh kiện","Hoàn tất","Đã hủy"});filterStatus.SelectedIndex=0;filterStatus.SelectedIndexChanged+=(_,__)=>RefreshView();searchPanel.Controls.Add(filterStatus);
        
        list.Controls.Add(searchPanel,0,0);
        grid.Dock=DockStyle.Fill;grid.ReadOnly=true;grid.AllowUserToAddRows=false;grid.AutoGenerateColumns=false;grid.SelectionMode=DataGridViewSelectionMode.FullRowSelect;grid.MultiSelect=false;grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill;grid.RowHeadersVisible=false;grid.SelectionChanged+=(_,__)=>LoadSelected();
        AddColumn("Mã đơn","MaDon");
        AddColumn("Tên khách hàng","TenKhachHang");
        AddColumn("Số điện thoại","SoDienThoai");
        AddColumn("Loại thiết bị","LoaiThietBi");
        AddColumn("Mô tả lỗi","MoTaLoi");
        AddColumn("Ngày tiếp nhận","NgayTiepNhan", "dd/MM/yyyy");
        AddColumn("Chi phí dự kiến","ChiPhiDuKien", "N0");
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
        if(string.IsNullOrWhiteSpace(txtName.Text)){errors.SetError(txtName,"Bắt buộc tên khách hàng.");ok=false;}if(!System.Text.RegularExpressions.Regex.IsMatch(txtPhone.Text.Trim(),@"^(0\d{9,10}|\+84\d{9,10})$")){errors.SetError(txtPhone,"Số điện thoại không hợp lệ.");ok=false;}if(string.IsNullOrWhiteSpace(txtDescription.Text)){errors.SetError(txtDescription,"Bắt buộc mô tả lỗi.");ok=false;}
        return ok;
    }
    private string GetId(RepairOrder x) => x.MaDon;
    private void AddItem(){if(!ValidateInput())return;items.Add(new RepairOrder{MaDon=txtId.Text.Trim(),TenKhachHang=txtName.Text.Trim(),SoDienThoai=txtPhone.Text.Trim(),LoaiThietBi=cboDevice.Text,MoTaLoi=txtDescription.Text.Trim(),NgayTiepNhan=dtMain.Value.Date,ChiPhiDuKien=numCost.Value,TrangThai=cboStatus.Text,DuongDanAnh=imagePath});RefreshView();ClearInputs();}
    private void UpdateItem(){if(selected==null){MessageBox.Show("Hãy chọn bản ghi cần cập nhật.");return;}if(!ValidateInput())return;var a=selected;a.MaDon=txtId.Text.Trim();a.TenKhachHang=txtName.Text.Trim();a.SoDienThoai=txtPhone.Text.Trim();a.LoaiThietBi=cboDevice.Text;a.MoTaLoi=txtDescription.Text.Trim();a.NgayTiepNhan=dtMain.Value.Date;a.ChiPhiDuKien=numCost.Value;a.TrangThai=cboStatus.Text;a.DuongDanAnh=imagePath;source.ResetBindings(false);RefreshView();}
    private void DeleteItem(){if(selected==null){MessageBox.Show("Hãy chọn bản ghi cần xóa.");return;}if(MessageBox.Show("Bạn có chắc muốn xóa?","Xác nhận",MessageBoxButtons.YesNo,MessageBoxIcon.Warning)!=DialogResult.Yes)return;items.Remove(selected);selected=null;ClearInputs();RefreshView();}
    private void ClearInputs(){selected=null;txtId.Clear();txtName.Clear();txtPhone.Clear();txtDescription.Clear();cboDevice.SelectedIndex=0;cboStatus.SelectedIndex=0;dtMain.Value=DateTime.Today;numCost.Value=0;imagePath="";pic.Image=null;errors.Clear();grid.ClearSelection();}
    private void LoadSelected(){if(grid.CurrentRow?.DataBoundItem is not RepairOrder a)return;selected=a;txtId.Text=a.MaDon;txtName.Text=a.TenKhachHang;txtPhone.Text=a.SoDienThoai;cboDevice.Text=a.LoaiThietBi;txtDescription.Text=a.MoTaLoi;dtMain.Value=a.NgayTiepNhan;numCost.Value=Math.Max(0,a.ChiPhiDuKien);cboStatus.Text=a.TrangThai;imagePath=a.DuongDanAnh;pic.Image=null;if(File.Exists(imagePath))try{pic.Image=Image.FromFile(imagePath);}catch{}}
    private void RefreshView(){var q=txtSearch.Text.Trim();var view=items.Where(x=>(x.MaDon.Contains(q,StringComparison.OrdinalIgnoreCase)||x.TenKhachHang.Contains(q,StringComparison.OrdinalIgnoreCase)||x.SoDienThoai.Contains(q,StringComparison.OrdinalIgnoreCase))&&((filterStatus.SelectedIndex==0||x.TrangThai==filterStatus.Text))).ToList();viewSource.DataSource=view;lblStatus.Text=$"Số đơn: {view.Count} | Chi phí chưa hoàn tất: {view.Where(x=>x.TrangThai!="Hoàn tất"&&x.TrangThai!="Đã hủy").Sum(x=>x.ChiPhiDuKien):N0} đ";}
    private void ExportCsv(){using var d=new SaveFileDialog{Filter="CSV (*.csv)|*.csv",FileName="RepairServiceTracker.csv"};if(d.ShowDialog()!=DialogResult.OK)return;var q=txtSearch.Text.Trim();var view=items.Where(x=>(x.MaDon.Contains(q,StringComparison.OrdinalIgnoreCase)||x.TenKhachHang.Contains(q,StringComparison.OrdinalIgnoreCase)||x.SoDienThoai.Contains(q,StringComparison.OrdinalIgnoreCase))&&((filterStatus.SelectedIndex==0||x.TrangThai==filterStatus.Text))).ToList();using var w=new StreamWriter(d.FileName,false,new UTF8Encoding(true));w.WriteLine(string.Join(",",new string[]{"Mã đơn","Tên khách hàng","Số điện thoại","Loại thiết bị","Mô tả lỗi","Ngày tiếp nhận","Chi phí dự kiến","Trạng thái"}.Select(Csv)));foreach(var a in view)w.WriteLine(string.Join(",",new[]{a.MaDon,a.TenKhachHang,a.SoDienThoai,a.LoaiThietBi,a.MoTaLoi,a.NgayTiepNhan.ToString("yyyy-MM-dd"),a.ChiPhiDuKien.ToString("N0"),a.TrangThai}.Select(Csv)));MessageBox.Show("Xuất CSV thành công.");}
    private static string Csv(string s) => "\"" + (s ?? "").Replace("\"", "\"\"") + "\"";
    private void ChooseImage(){using var d=new OpenFileDialog{Filter="Ảnh|*.png;*.jpg;*.jpeg;*.bmp;*.gif"};if(d.ShowDialog()!=DialogResult.OK)return;imagePath=d.FileName;try{using var fs=new FileStream(imagePath,FileMode.Open,FileAccess.Read);using var img=Image.FromStream(fs);pic.Image=new Bitmap(img);}catch{MessageBox.Show("Không đọc được ảnh.");}}
}
