using System.Data;
using System.Drawing;
using System.Windows.Forms;
using ProcessadorProjetos.Services;

namespace ProcessadorProjetos;

public partial class MainWindow : Form
{
    private string _caminhoArquivo = "";
    private string _diretorioSaida = "";
    private DataTable _dados = new DataTable();

    public MainWindow()
    {
        InitializeComponent();
        SetupUI();
    }

    private void SetupUI()
    {
        this.Text = "Processador de Projetos";
        this.Size = new Size(650, 400);
        this.StartPosition = FormStartPosition.CenterScreen;

        var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
        this.Controls.Add(panel);

        var lblArq = new Label { Text = "Arquivo Excel:", Font = new Font("Segoe UI", 10, FontStyle.Bold), Top = 10, Left = 20 };
        panel.Controls.Add(lblArq);

        var txtArq = new TextBox { Width = 400, Top = 40, Left = 20 };
        txtArq.TextChanged += (s, e) => _caminhoArquivo = txtArq.Text;
        panel.Controls.Add(txtArq);

        var btnProc = new Button { Text = "Procurar", Width = 100, Top = 38, Left = 430 };
        btnProc.Click += (s, e) => {
            using var dialog = new OpenFileDialog { Filter = "Excel|*.xls;*.xlsx" };
            if (dialog.ShowDialog() == DialogResult.OK) {
                _caminhoArquivo = dialog.FileName;
                txtArq.Text = _caminhoArquivo;
            }
        };
        panel.Controls.Add(btnProc);

        var lblSaida = new Label { Text = "Diretório de Saída:", Font = new Font("Segoe UI", 10, FontStyle.Bold), Top = 80, Left = 20 };
        panel.Controls.Add(lblSaida);

        var txtSaida = new TextBox { Width = 400, Top = 110, Left = 20 };
        txtSaida.TextChanged += (s, e) => _diretorioSaida = txtSaida.Text;
        panel.Controls.Add(txtSaida);

        var btnSel = new Button { Text = "Selecionar", Width = 100, Top = 108, Left = 430 };
        btnSel.Click += (s, e) => {
            using var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK) {
                _diretorioSaida = dialog.SelectedPath;
                txtSaida.Text = _diretorioSaida;
            }
        };
        panel.Controls.Add(btnSel);

        var btnProcessar = new Button { Text = "Processar Excel", Width = 150, Top = 160, Left = 20, Height = 40 };
        btnProcessar.Click += ProcessarArquivo;
        panel.Controls.Add(btnProcessar);

        var btnRelatorio = new Button { Text = "Gerar PDF", Width = 150, Top = 160, Left = 180, Height = 40 };
        btnRelatorio.Click += GerarRelatorio;
        panel.Controls.Add(btnRelatorio);
    }

    private void ProcessarArquivo(object? sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(_caminhoArquivo) || string.IsNullOrEmpty(_diretorioSaida))
        {
            MessageBox.Show("Selecione arquivo e diretório!", "Erro");
            return;
        }
        try
        {
            _dados = ExcelService.LerPlanilha(_caminhoArquivo);
            var caminhoSaida = Path.Combine(_diretorioSaida, $"Processado_{DateTime.Now:ddMMyyyy_HHmmss}.xlsx");
            ExcelService.SalvarPlanilha(_dados, caminhoSaida);
            MessageBox.Show($"Sucesso!\n{caminhoSaida}", "Concluído");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Erro: {ex.Message}", "Erro");
        }
    }

    private void GerarRelatorio(object? sender, EventArgs e)
    {
        if (_dados.Rows.Count == 0 || string.IsNullOrEmpty(_diretorioSaida))
        {
            MessageBox.Show("Processe o arquivo primeiro!", "Erro");
            return;
        }
        try
        {
            var caminhoPdf = Path.Combine(_diretorioSaida, $"Relatorio_{DateTime.Now:ddMMyyyy_HHmmss}.pdf");
            PdfService.GerarRelatorio(_dados, caminhoPdf);
            MessageBox.Show($"Relatório gerado!\n{caminhoPdf}", "Sucesso");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Erro: {ex.Message}", "Erro");
        }
    }
}
