using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;

namespace PerformaTools
{
    public partial class regraGeral : Form
    {
        // --- Nomes das colunas ---
        private const string ColNomeFluxo = "colNomeFluxo";
        private const string ColAtivacao = "colAtivacao";
        private const string ColDetalharRegra = "colDetalharRegra";
        private const string ColResponsabilidade = "colResponsabilidade";
        private const string ColPerfisAbrangentes = "colPerfisAbrangentes";
        private const string ColUsuarioEspecifico = "colUsuarioEspecifico";
        private const string ColQtdDias = "colQtdDias";

        // --- UI ---
        private const string DISABLED_DASH = "—";

        // Botão embutido
        private const int BTN_WIDTH = 22;
        private const int BTN_PAD = 4;
        private const int TEXT_RIGHT_PAD = 4;

        // Cores dos “cards”
        private static readonly Color CardBlueBack = Color.FromArgb(232, 241, 255);
        private static readonly Color CardBlueBorder = Color.FromArgb(150, 180, 220);
        private static readonly Color CardYellowBack = Color.FromArgb(255, 249, 219);
        private static readonly Color CardYellowBorder = Color.FromArgb(255, 224, 138);

        // Layout raiz
        private TableLayoutPanel _layout = null!;
        private Panel _cardInstrucoes = null!;
        private Panel _cardProximos = null!;
        private Panel _panelGrid = null!;
        private Panel _panelButtons = null!;
        private FlowLayoutPanel _btnLeft = null!;
        private FlowLayoutPanel _btnRight = null!;

        private Button _btnLimpar = null!;
        private Button _btnGerarExcel = null!;
        private Button _btnFechar = null!;

        private DataGridView grid = null!;

        // Debounce
        private bool _openingPerfis = false;
        private DateTime _lastPerfisOpen = DateTime.MinValue;

        private bool _openingUsuario = false;
        private DateTime _lastUsuarioOpen = DateTime.MinValue;

        private bool _openingDias = false;
        private DateTime _lastDiasOpen = DateTime.MinValue;

        // --- Opções canônicas ---
        private static readonly string OP_ADV = "Advogado Responsável do Processo";
        private static readonly string OP_DIST = "Distribuição Igualitária Dia";
        private static readonly string OP_CONC = "Responsável conclusão do Prazo/PA";
        private static readonly string OP_PRAZO_PA = "Responsável do Prazo/PA";
        private static readonly string OP_REVI = "Responável pela Revisão";
        private static readonly string OP_SELEC = "Selecionar Responsável";
        private static readonly string OP_CAD = "Usuário Cadastro";
        private static readonly string OP_USU_ENC = "Usuário Encaminhamento Acordo";
        private static readonly string OP_USER = "Usuário Específico";

        // Fluxo -> opções de responsabilidade
        private static readonly Dictionary<string, string[]> FLUXO_RESPONSABILIDADE = new()
        {
            { "Acordo", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Acordo Pós Sentença", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Aprovação Alteração Prazo", new[] { OP_ADV, OP_DIST, OP_USER, OP_PRAZO_PA } },
            { "Aprovação Cancelamento de Obrigação Processo", new[] { OP_ADV, OP_DIST, OP_USER, OP_PRAZO_PA } },
            { "Aprovação Cancelamento Prazo", new[] { OP_ADV, OP_DIST, OP_USER, OP_PRAZO_PA } },
            { "Aud. Solicitação Audiência", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER, OP_PRAZO_PA } },
            { "Audiência Cancelada - PA", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER, OP_PRAZO_PA } },
            { "Audiência Designada - PA", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER, OP_PRAZO_PA } },
            { "Auditoria de Solicitação Diligência", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER, OP_PRAZO_PA } },
            { "Confirmação de Patrocínio", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Encerramento", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "FollowupContraProposta", new[] { OP_ADV, OP_DIST, OP_USER, OP_PRAZO_PA, OP_USU_ENC } },
            { "Garantia", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Intimação", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Leitura de Sentença - PA", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER, OP_PRAZO_PA } },
            { "Notificação Conclusão Prazo", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER, OP_PRAZO_PA } },
            { "Notificação de Remoção da Negociação", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Obrigação", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER, OP_PRAZO_PA } },
            { "Obrigação Ilíquida", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER, OP_PRAZO_PA } },
            { "PA", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "PA Recurso Parte Adversa", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER, OP_PRAZO_PA } },
            { "Prazo", new[] { OP_ADV, OP_DIST, OP_USER, OP_SELEC } },
            { "Primeira Audiência", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER, OP_PRAZO_PA } },
            { "Processo Sem Movimentação", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Protocolo Físico", new[] { OP_ADV, OP_DIST, OP_USER, OP_CONC, OP_REVI, OP_PRAZO_PA } },
            { "Protocolo Virtual", new[] { OP_ADV, OP_DIST, OP_USER, OP_CONC, OP_REVI, OP_PRAZO_PA } },
            { "Publicação",  new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Publicação Responsabilidade Auxiliar",  new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Redesignação - PA", new[] { OP_ADV, OP_DIST, OP_USER, OP_REVI, OP_PRAZO_PA } },
            { "Repasse", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Responsável Análise de Recurso", new[] { OP_ADV, OP_DIST, OP_USER, OP_REVI, OP_PRAZO_PA } },
            { "Revisão de Intimação", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Revisão de PA", new[] { OP_ADV, OP_DIST, OP_USER, OP_REVI, OP_PRAZO_PA } },
            { "Revisão de Prazo", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER, OP_SELEC, OP_PRAZO_PA } },
            { "Revisão de Processo", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Revisão de Publicação", new[] { OP_ADV, OP_DIST, OP_USER } },
            { "Sessão de Julgamento - PA", new[] { OP_ADV, OP_DIST, OP_CAD, OP_USER } },
            { "Subsídio", new[] { OP_ADV, OP_DIST, OP_USER, OP_PRAZO_PA } },
            { "Validação Cliente", new[] { OP_ADV, OP_DIST, OP_USER, OP_REVI, OP_PRAZO_PA } },
        };

        // Perfis (lista multi)
        private static readonly string[] PERFIS_OPCOES = new[]
        {
            "Assistente Coordenação","Assistente Coordenação Consultivo","Assistente LGPD Consultivo",
            "Assistente Núcleo Jurídico","Assistente QCA","Assistente Unificado","Assistente Unificado Consultivo",
            "Backoffice - Cadastro","Backoffice Operacional",
            "Controladoria - Núcleo de Serviços","Controladoria - Núcleo de Serviços Financeiro",
            "Controladoria - Núcleo Performa Bots","Controladoria - Núcleo Tratamento","Controladoria - Pauta QCA",
            "Controladoria - Reports","Controladoria - Revisão de Cadastro","Controladoria – Supervisão Núcleo de Serviços",
            "Controladoria - Supervisor de Pautas e Serviços","Controladoria QCA","Coordenador",
            "Coordenador - Pautas e Serviços","Coordenador - Projeto Impulsionamento","Coordenador Consultivo",
            "Coordenador Controladoria Jurídica","Coordenador de Treinamento","Coordenador Estratégico",
            "Coordenador Núcleo Agendamento","Coordenador Núcleo Controller","Coordenador Núcleo Cumprimento",
            "Coordenador Núcleo de Acordo com Cumprimento","Coordenador Núcleo Encerramento","Coordenador Núcleo Execução",
            "Coordenador Núcleo Pauta de Audiência","Coordenador Núcleo Protocolo","Coordenador Núcleo Redação",
            "Coordenador Núcleo Subsídio","Coordenador Núcleo Subsídio com Acordo","Coordenador Redação Middle",
            "Coordenador Unificado","Estagiário - Pautas e Serviços","Estagiário Cumprimento",
            "Estagiário Cumprimento OBF","Estagiário Cumprimento OBP","Estagiário Nova Demanda",
            "Estagiário Núcleo de Custas","Estagiário Núcleo Estratégico","Estagiário Núcleo Execução",
            "Estagiário Núcleo Redação","Estagiário Núcleo Subsídio","Estagiário Pós-Sentença",
            "Estagiário Redação Pós-Sentença","Estagiário Unificado","Estagiário Unificado Consultivo",
            "Gestor","Gestor Consultivo","Gestor de Unidade","Gestor Redação Middle","Negociador NAC",
            "Núcleo Agendamento","Núcleo Controller","Núcleo Cumprimento","Núcleo Cumprimento OBF","Núcleo Cumprimento OBP",
            "Núcleo de Acordo","Núcleo de Acordo com Cumprimento","Núcleo de Acordo Subsídios Cumprimento",
            "Núcleo de Cálculos Judiciais","Núcleo de Citação","Núcleo de Custas","Núcleo de Impulsionamento",
            "Núcleo de Penhora","Núcleo de Pesquisa Patrimonial","Núcleo de Suporte - Unidades QCA",
            "Núcleo Encerramento","Núcleo Estratégico","Núcleo Execução","Núcleo Habilitação","Núcleo Laudo Técnico",
            "Núcleo Nova Demanda","Núcleo Pauta de Audiência","Núcleo Performa Bots","Núcleo Pós-Sentença",
            "Núcleo Protocolo","Núcleo Protocolo com Cumprimento","Núcleo Protocolo Habilitação",
            "Núcleo Redação","Núcleo Redação Middle","Núcleo Redação Pós-Sentença","Núcleo Subsídio","Núcleo Suporte",
            "Núcleo Suporte - Cumprimento e Encerramento","Núcleo Suporte Operacional","Núcleo Tratamento Nova Demanda",
            "Prestador Externo","Prestador Externo - Encerramento","Prestador Externo - Execução","Prestador Externo - Nova Demanda",
            "Prestador Externo – Recursal","Prestador Externo - Redação Middle","Prestador Externo - Revisão de Base",
            "Prestador Externo Acordo","Supervisor - Redação Middle","Supervisor NAC","Supervisor NAC Virtual",
            "Supervisor Núcleo Controller","Supervisor Núcleo Controller Unificado","Supervisor Núcleo Cumprimento",
            "Supervisor Núcleo de Acordo com Cumprimento","Supervisor Núcleo Estratégico","Supervisor Núcleo Execução",
            "Supervisor Núcleo Pós-Sentença","Supervisor Núcleo Pré-Sentença","Supervisor Núcleo Protocolo",
            "Supervisor Núcleo Redação","Supervisor Pauta de Audiência"
        };

        public regraGeral()
        {
            MontarUIBasica();
            DefinirColunas();
            CarregarFluxosFixos(new[]
            {
                "Acordo","Acordo Pós Sentença","Aprovação Alteração Prazo",
                "Aprovação Cancelamento de Obrigação Processo","Aprovação Cancelamento Prazo",
                "Aud. Solicitação Audiência","Audiência Cancelada - PA","Audiência Designada - PA",
                "Auditoria de Solicitação Diligência","Confirmação de Patrocínio","Encerramento",
                "FollowupContraProposta","Garantia","Intimação","Leitura de Sentença - PA",
                "Notificação Conclusão Prazo","Notificação de Remoção da Negociação","Obrigação",
                "Obrigação Ilíquida","PA","PA Recurso Parte Adversa","Prazo","Primeira Audiência",
                "Processo Sem Movimentação","Protocolo Físico","Protocolo Virtual","Publicação",
                "Publicação Responsabilidade Auxiliar","Redesignação - PA","Repasse",
                "Responsável Análise de Recurso","Revisão de Intimação","Revisão de PA",
                "Revisão de Prazo","Revisão de Processo","Revisão de Publicação","Sessão de Julgamento - PA",
                "Subsídio","Validação Cliente"
            });

            // Combos 1 clique
            grid.CellClick += Grid_CellClick;
            grid.CellEnter += Grid_CellEnter;
            grid.EditingControlShowing += Grid_EditingControlShowing;
            grid.CellValueChanged += Grid_CellValueChanged;

            // PERFIS
            grid.CellClick += Grid_CellClick_Perfis;
            grid.CellMouseClick += Grid_CellMouseClick_PerfisButton;
            grid.CellPainting += Grid_CellPainting_PerfisButton;

            // USUÁRIO
            grid.CellClick += Grid_CellClick_Usuario;
            grid.CellMouseClick += Grid_CellMouseClick_UsuarioButton;
            grid.CellPainting += Grid_CellPainting_UsuarioButton;

            // DIAS
            grid.CellClick += Grid_CellClick_Dias;
            grid.CellMouseClick += Grid_CellMouseClick_DiasButton;
            grid.CellPainting += Grid_CellPainting_DiasButton;

            AplicarOpcoesResponsabilidadePorFluxo();
            AtualizarHabilitacaoParaTodas();

            // <<< RESPONSIVIDADE: calcular mínimos por conteúdo e preencher tudo >>>
            AplicarAutoLarguraResponsiva();

            // se redimensionar janela, Fill já acompanha; mas se quiser reavaliar após temas/fontes:
            this.FontChanged += (s, e) => AplicarAutoLarguraResponsiva();
        }

        // =====================  UI / Layout  =====================
        private void MontarUIBasica()
        {
            Text = "PerformaTools - Fluxos";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 1220; Height = 760;

            _layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(16),
            };
            _layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(_layout);

            _cardInstrucoes = CreateCard(
                "Instruções",
                "Preencha apenas os campos obrigatórios abaixo para gerar o relatório de alteração das regras gerais. " +
                "O sistema realizará validações automáticas para garantir a precisão da sua solicitação.",
                CardBlueBack, CardBlueBorder
            );
            _layout.Controls.Add(_cardInstrucoes, 0, 0);

            _cardProximos = CreateCard(
                "Próximos Passos",
                "• Preencha apenas os fluxos que haverá mudança.\n" +
                "• Após o preenchimento, clique em “Gerar Excel” para criar o arquivo no formato padronizado.\n" +
                "• Salve o arquivo Excel gerado em seu computador.\n" +
                "• Acesse o sistema de chamados GLPI.\n" +
                "• Anexe o arquivo Excel ao criar um novo chamado na categoria “Alteração de Configuração de Célula >> Regra Geral”.",
                CardYellowBack, CardYellowBorder
            );
            _layout.Controls.Add(_cardProximos, 0, 1);

            _panelGrid = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 8, 0, 8) };
            _layout.Controls.Add(_panelGrid, 0, 2);

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoGenerateColumns = false,
                EditMode = DataGridViewEditMode.EditOnEnter,
                // IMPORTANTE: responsivo com Fill (larguras mínimas serão setadas depois)
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            _panelGrid.Controls.Add(grid);

            // Cabeçalho azul
            var azulCabecalho = Color.FromArgb(0, 32, 96);
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = azulCabecalho;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = azulCabecalho;
            grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = Color.White;
            grid.ColumnHeadersHeight = 34;

            // Zebra + seleção
            grid.BackgroundColor = Color.White;
            grid.RowsDefaultCellStyle.BackColor = Color.White;
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            var azulSelecao = Color.FromArgb(0, 32, 96);
            grid.DefaultCellStyle.SelectionBackColor = azulSelecao;
            grid.DefaultCellStyle.SelectionForeColor = Color.White;
            grid.AlternatingRowsDefaultCellStyle.SelectionBackColor = azulSelecao;
            grid.AlternatingRowsDefaultCellStyle.SelectionForeColor = Color.White;

            grid.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (grid.IsCurrentCellDirty)
                    grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            grid.DataError += (s, e) => { e.ThrowException = false; };

            // Botões
            _panelButtons = new Panel { Dock = DockStyle.Fill, Height = 56 };
            _layout.Controls.Add(_panelButtons, 0, 3);

            _btnLeft = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Padding = new Padding(0, 8, 0, 8)
            };
            _panelButtons.Controls.Add(_btnLeft);

            _btnRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(0, 8, 0, 8)
            };
            _panelButtons.Controls.Add(_btnRight);

            _btnLimpar = MakeButton("Limpar Preenchimento", Color.Gainsboro, Color.Black);
            _btnLimpar.Click += (s, e) => { LimparPreenchimento(); AplicarAutoLarguraResponsiva(); };
            _btnLeft.Controls.Add(_btnLimpar);

            _btnGerarExcel = MakeButton("Gerar Excel", Color.FromArgb(46, 125, 50), Color.White);
            _btnGerarExcel.Click += (s, e) => ExportadorExcel.ExportRegraGeral(grid, this);

            _btnFechar = MakeButton("Fechar", Color.White, Color.Black, Color.Silver);
            _btnFechar.Click += (s, e) => Close();

            // IMPORTANTE: como o fluxo é RightToLeft, adicionamos nesta ordem
            _btnRight.Controls.Add(_btnFechar);  // fica colado à borda direita
            _btnRight.Controls.Add(_btnGerarExcel);      // fica imediatamente à esquerda do "Gerar Excel"
        }

        private static Button MakeButton(string text, Color back, Color fore, Color? border = null)
        {
            var b = new Button
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = back,
                ForeColor = fore,
                FlatStyle = FlatStyle.Flat,
                Padding = new Padding(14, 6, 14, 6),
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand,
                Margin = new Padding(6, 0, 6, 0)
            };
            b.FlatAppearance.BorderSize = 1;
            b.FlatAppearance.BorderColor = border ?? back;
            return b;
        }

        private static Panel CreateCard(string titulo, string texto, Color backColor, Color borderColor)
        {
            var card = new Panel
            {
                BackColor = backColor,
                Padding = new Padding(14),
                Margin = new Padding(0, 0, 0, 10),
                AutoSize = true
            };

            var lblTitle = new Label
            {
                Text = titulo,
                AutoSize = true,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
                ForeColor = Color.FromArgb(35, 45, 65)
            };

            var lblBody = new Label
            {
                Text = texto,
                AutoSize = true,
                MaximumSize = new Size(1100, 0)
            };

            card.Controls.Add(lblBody);
            card.Controls.Add(lblTitle);
            lblTitle.Location = new Point(6, 6);
            lblBody.Location = new Point(6, lblTitle.Bottom + 8);

            card.Paint += (s, e) =>
            {
                var g = e.Graphics;
                if (g == null) return;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using var pen = new Pen(borderColor, 1f);
                using var path = RoundedRect(new Rectangle(0, 0, card.Width - 1, card.Height - 1), 8);
                g.DrawPath(pen, path);
            };

            card.SizeChanged += (s, e) => card.Invalidate();
            return card;
        }

        private static GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            int d = radius * 2;
            var path = new GraphicsPath();
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        // =====================  GRID  =====================
        private void DefinirColunas()
        {
            var cNome = new DataGridViewTextBoxColumn
            {
                Name = ColNomeFluxo,
                HeaderText = "Nome Fluxo",
                ReadOnly = true
            };
            grid.Columns.Add(cNome);

            var cAtivacao = new DataGridViewComboBoxColumn
            {
                Name = ColAtivacao,
                HeaderText = "Ativar/Inativar",
                FlatStyle = FlatStyle.Flat,
                Visible = true
            };
            cAtivacao.Items.AddRange("Selecione", "Ativar", "Inativar");
            grid.Columns.Add(cAtivacao);

            var estiloDetalhar = new DataGridViewCellStyle
            {
                BackColor = Color.Gainsboro,
                ForeColor = Color.DimGray
            };
            var cDetalhar = new DataGridViewTextBoxColumn
            {
                Name = ColDetalharRegra,
                HeaderText = "Detalhar Regra",
                ReadOnly = true,
                Visible = true,
                DefaultCellStyle = estiloDetalhar
            };
            grid.Columns.Add(cDetalhar);

            var cResp = new DataGridViewComboBoxColumn
            {
                Name = ColResponsabilidade,
                HeaderText = "Responsabilidade",
                FlatStyle = FlatStyle.Flat,
                Visible = true
            };
            cResp.Items.AddRange("Selecione", OP_DIST, OP_USER);
            grid.Columns.Add(cResp);

            var cPerfis = new DataGridViewTextBoxColumn
            {
                Name = ColPerfisAbrangentes,
                HeaderText = "Perfis acesso abrangentes",
                Visible = true,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Padding = new Padding(BTN_PAD + BTN_WIDTH + BTN_PAD, 0, TEXT_RIGHT_PAD, 0)
                }
            };
            grid.Columns.Add(cPerfis);

            var cUsuario = new DataGridViewTextBoxColumn
            {
                Name = ColUsuarioEspecifico,
                HeaderText = "Caso usuário específico",
                Visible = true,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Padding = new Padding(BTN_PAD + BTN_WIDTH + BTN_PAD, 0, TEXT_RIGHT_PAD, 0)
                }
            };
            grid.Columns.Add(cUsuario);

            var cDias = new DataGridViewTextBoxColumn
            {
                Name = ColQtdDias,
                HeaderText = "Quantidade de Dias Expectativa",
                Visible = true,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    Padding = new Padding(BTN_PAD + BTN_WIDTH + BTN_PAD, 0, TEXT_RIGHT_PAD, 0)
                }
            };
            grid.Columns.Add(cDias);
        }

        private void CarregarFluxosFixos(string[] fluxos)
        {
            grid.Rows.Clear();
            foreach (var nome in fluxos)
            {
                int idx = grid.Rows.Add();
                var row = grid.Rows[idx];

                row.Cells[ColNomeFluxo].Value = nome;
                row.Cells[ColAtivacao].Value = "Selecione";
                row.Cells[ColDetalharRegra].Value = "Não";
                row.Cells[ColResponsabilidade].Value = "Selecione";
                row.Cells[ColPerfisAbrangentes].Value = "";
                row.Cells[ColUsuarioEspecifico].Value = "";
                row.Cells[ColQtdDias].Value = "";
            }
        }

        private void AplicarOpcoesResponsabilidadePorFluxo()
        {
            foreach (DataGridViewRow row in grid.Rows)
            {
                var fluxo = Convert.ToString(row.Cells[ColNomeFluxo].Value)?.Trim() ?? "";
                if (row.Cells[ColResponsabilidade] is not DataGridViewComboBoxCell cell) continue;

                cell.Items.Clear();
                cell.Items.Add("Selecione");

                if (FLUXO_RESPONSABILIDADE.TryGetValue(fluxo, out var opcoes))
                    foreach (var op in opcoes) cell.Items.Add(op);
                else
                {
                    cell.Items.Add(OP_DIST);
                    cell.Items.Add(OP_USER);
                }

                cell.Value = "Selecione";
            }
        }

        // ===== Combos 1 clique =====
        private void Grid_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var colName = grid.Columns[e.ColumnIndex].Name;
            if (colName == ColAtivacao || colName == ColResponsabilidade)
            {
                grid.BeginEdit(true);
                BeginInvoke(new Action(() =>
                {
                    if (grid.EditingControl is ComboBox cb) cb.DroppedDown = true;
                }));
            }
        }
        private void Grid_CellEnter(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var colName = grid.Columns[e.ColumnIndex].Name;
            if (colName == ColAtivacao || colName == ColResponsabilidade)
            {
                grid.BeginEdit(true);
                BeginInvoke(new Action(() =>
                {
                    if (grid.EditingControl is ComboBox cb) cb.DroppedDown = true;
                }));
            }
        }
        private void Grid_EditingControlShowing(object? sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is ComboBox cb)
            {
                cb.DropDownStyle = ComboBoxStyle.DropDownList;
                cb.IntegralHeight = false;
                cb.MaxDropDownItems = 12;
            }
        }

        // ===== Regras reativas =====
        private void Grid_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var col = grid.Columns[e.ColumnIndex].Name;

            if (col == ColResponsabilidade)
            {
                AtualizarHabilitacaoLinha(e.RowIndex);

                var resp = Convert.ToString(grid.Rows[e.RowIndex].Cells[ColResponsabilidade].Value) ?? "";
                if (!string.Equals(resp, OP_DIST, StringComparison.OrdinalIgnoreCase))
                    grid.Rows[e.RowIndex].Cells[ColPerfisAbrangentes].Value = "";

                if (!string.Equals(resp, OP_USER, StringComparison.OrdinalIgnoreCase))
                    grid.Rows[e.RowIndex].Cells[ColUsuarioEspecifico].Value = "";
            }
        }

        // ===== PERFIS =====
        private bool IsPerfisEnabled(int rowIndex)
        {
            var resp = Convert.ToString(grid.Rows[rowIndex].Cells[ColResponsabilidade].Value) ?? "";
            return string.Equals(resp, OP_DIST, StringComparison.OrdinalIgnoreCase);
        }
        private void Grid_CellClick_Perfis(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (grid.Columns[e.ColumnIndex].Name != ColPerfisAbrangentes) return;
            if (!IsPerfisEnabled(e.RowIndex)) return;

            if (_openingPerfis) return;
            if ((DateTime.UtcNow - _lastPerfisOpen).TotalMilliseconds < 250) return;
            AbrirPerfisPicker(e.RowIndex);
        }
        private void Grid_CellMouseClick_PerfisButton(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (grid.Columns[e.ColumnIndex].Name != ColPerfisAbrangentes) return;
            if (!IsPerfisEnabled(e.RowIndex)) return;

            var btnRect = new Rectangle(BTN_PAD, (grid.Rows[e.RowIndex].Height - 18) / 2, BTN_WIDTH, 18);
            if (!btnRect.Contains(e.Location)) return;

            if (_openingPerfis) return;
            if ((DateTime.UtcNow - _lastPerfisOpen).TotalMilliseconds < 250) return;
            AbrirPerfisPicker(e.RowIndex);
        }
        private void Grid_CellPainting_PerfisButton(object? sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (grid.Columns[e.ColumnIndex].Name != ColPerfisAbrangentes) return;
            if (!IsPerfisEnabled(e.RowIndex)) return;

            e.Paint(e.CellBounds, DataGridViewPaintParts.All);

            var g = e.Graphics;
            if (g == null) { e.Handled = true; return; }
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            int btnH = Math.Min(18, e.CellBounds.Height - 4);
            var btnRect = new Rectangle(
                e.CellBounds.Left + BTN_PAD,
                e.CellBounds.Top + (e.CellBounds.Height - btnH) / 2,
                BTN_WIDTH,
                btnH
            );
            using var back = new SolidBrush(Color.LightGray);
            using var border = new Pen(Color.Gray);
            using var txtBrush = new SolidBrush(Color.Black);
            g.FillRectangle(back, btnRect);
            g.DrawRectangle(border, btnRect);
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString("⋯", grid.Font, txtBrush, btnRect, sf);

            e.Handled = true;
        }
        private void AbrirPerfisPicker(int rowIndex)
        {
            _openingPerfis = true;
            try
            {
                var cell = grid.Rows[rowIndex].Cells[ColPerfisAbrangentes];

                var atual = Convert.ToString(cell.Value) ?? "";
                if (IsPlaceholder(atual)) atual = "";
                var preSel = atual.Split(';').Select(s => s.Trim()).Where(s => s.Length > 0)
                                  .ToHashSet(StringComparer.OrdinalIgnoreCase);

                int colIndex = grid.Columns[ColPerfisAbrangentes]?.Index ?? -1;
                if (colIndex < 0) return;
                var rect = grid.GetCellDisplayRectangle(colIndex, rowIndex, true);
                var screenPoint = grid.PointToScreen(new Point(rect.Left, rect.Bottom));
                int ddWidth = Math.Max(rect.Width, 420);

                var escolhidos = PerfisDropDown.ShowAt(screenPoint, ddWidth, 340, PERFIS_OPCOES, preSel, this);
                if (escolhidos == null) return;

                if (escolhidos.Length == 0)
                    SetPerfisPlaceholder(rowIndex);
                else
                {
                    cell.Value = string.Join("; ", escolhidos);
                    SetPerfisNormalStyle(rowIndex);
                }

                if (grid.IsCurrentCellInEditMode) grid.EndEdit();
                grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                grid.UpdateCellValue(colIndex, rowIndex);
                grid.InvalidateCell(colIndex, rowIndex);
                grid.Refresh();

                // pode ter ficado mais largo:
                AplicarAutoLarguraResponsiva();
            }
            finally { _openingPerfis = false; _lastPerfisOpen = DateTime.UtcNow; }
        }

        // ===== USUÁRIO ESPECÍFICO =====
        private bool IsUsuarioEnabled(int rowIndex)
        {
            var resp = Convert.ToString(grid.Rows[rowIndex].Cells[ColResponsabilidade].Value) ?? "";
            return string.Equals(resp, OP_USER, StringComparison.OrdinalIgnoreCase);
        }
        private void Grid_CellClick_Usuario(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (grid.Columns[e.ColumnIndex].Name != ColUsuarioEspecifico) return;
            if (!IsUsuarioEnabled(e.RowIndex)) return;

            if (_openingUsuario) return;
            if ((DateTime.UtcNow - _lastUsuarioOpen).TotalMilliseconds < 250) return;
            AbrirUsuarioDialog(e.RowIndex);
        }
        private void Grid_CellMouseClick_UsuarioButton(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (grid.Columns[e.ColumnIndex].Name != ColUsuarioEspecifico) return;
            if (!IsUsuarioEnabled(e.RowIndex)) return;

            var btnRect = new Rectangle(BTN_PAD, (grid.Rows[e.RowIndex].Height - 18) / 2, BTN_WIDTH, 18);
            if (!btnRect.Contains(e.Location)) return;

            if (_openingUsuario) return;
            if ((DateTime.UtcNow - _lastUsuarioOpen).TotalMilliseconds < 250) return;
            AbrirUsuarioDialog(e.RowIndex);
        }
        private void Grid_CellPainting_UsuarioButton(object? sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (grid.Columns[e.ColumnIndex].Name != ColUsuarioEspecifico) return;
            if (!IsUsuarioEnabled(e.RowIndex)) return;

            e.Paint(e.CellBounds, DataGridViewPaintParts.All);

            var g = e.Graphics;
            if (g == null) { e.Handled = true; return; }
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            int btnH = Math.Min(18, e.CellBounds.Height - 4);
            var btnRect = new Rectangle(
                e.CellBounds.Left + BTN_PAD,
                e.CellBounds.Top + (e.CellBounds.Height - btnH) / 2,
                BTN_WIDTH,
                btnH
            );
            using var back = new SolidBrush(Color.LightGray);
            using var border = new Pen(Color.Gray);
            using var txtBrush = new SolidBrush(Color.Black);
            g.FillRectangle(back, btnRect);
            g.DrawRectangle(border, btnRect);
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString("⋯", grid.Font, txtBrush, btnRect, sf);

            e.Handled = true;
        }
        private void AbrirUsuarioDialog(int rowIndex)
        {
            _openingUsuario = true;
            try
            {
                int colIndex = grid.Columns[ColUsuarioEspecifico]?.Index ?? -1;
                if (colIndex < 0) return;

                var cell = grid.Rows[rowIndex].Cells[colIndex];
                string atual = Convert.ToString(cell.Value) ?? "";
                if (IsPlaceholder(atual)) atual = "";

                var rect = grid.GetCellDisplayRectangle(colIndex, rowIndex, true);
                int w = Math.Max(rect.Width, 420);
                int h = 150;
                var screenPoint = grid.PointToScreen(new Point(rect.Left, rect.Bottom));

                var novoValor = UsuarioInputDialog.ShowAt(
                    screenPoint, w, h, "Informar usuário específico", atual, this);

                if (novoValor == null) return;

                novoValor = (novoValor ?? "").Trim();
                cell.Value = string.IsNullOrEmpty(novoValor) ? DISABLED_DASH : novoValor;

                if (grid.IsCurrentCellInEditMode) grid.EndEdit();
                grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                grid.UpdateCellValue(colIndex, rowIndex);
                grid.InvalidateCell(colIndex, rowIndex);
                grid.Refresh();

                if (string.IsNullOrEmpty(novoValor))
                    SetUsuarioEnabledPlaceholder(rowIndex);
                else
                    SetUsuarioEnabledFilled(rowIndex);

                AplicarAutoLarguraResponsiva();
            }
            finally { _openingUsuario = false; _lastUsuarioOpen = DateTime.UtcNow; }
        }

        // ===== DIAS (Encerramento) =====
        private bool IsDiasEnabled(int rowIndex)
        {
            var fluxo = Convert.ToString(grid.Rows[rowIndex].Cells[ColNomeFluxo].Value) ?? "";
            return string.Equals(fluxo, "Encerramento", StringComparison.OrdinalIgnoreCase);
        }
        private void Grid_CellClick_Dias(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (grid.Columns[e.ColumnIndex].Name != ColQtdDias) return;
            if (!IsDiasEnabled(e.RowIndex)) return;

            if (_openingDias) return;
            if ((DateTime.UtcNow - _lastDiasOpen).TotalMilliseconds < 250) return;
            AbrirDiasDialog(e.RowIndex);
        }
        private void Grid_CellMouseClick_DiasButton(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (grid.Columns[e.ColumnIndex].Name != ColQtdDias) return;
            if (!IsDiasEnabled(e.RowIndex)) return;

            var btnRect = new Rectangle(BTN_PAD, (grid.Rows[e.RowIndex].Height - 18) / 2, BTN_WIDTH, 18);
            if (!btnRect.Contains(e.Location)) return;

            if (_openingDias) return;
            if ((DateTime.UtcNow - _lastDiasOpen).TotalMilliseconds < 250) return;
            AbrirDiasDialog(e.RowIndex);
        }
        private void Grid_CellPainting_DiasButton(object? sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (grid.Columns[e.ColumnIndex].Name != ColQtdDias) return;
            if (!IsDiasEnabled(e.RowIndex)) return;

            e.Paint(e.CellBounds, DataGridViewPaintParts.All);

            var g = e.Graphics;
            if (g == null) { e.Handled = true; return; }
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            int btnH = Math.Min(18, e.CellBounds.Height - 4);
            var btnRect = new Rectangle(
                e.CellBounds.Left + BTN_PAD,
                e.CellBounds.Top + (e.CellBounds.Height - btnH) / 2,
                BTN_WIDTH,
                btnH
            );
            using var back = new SolidBrush(Color.LightGray);
            using var border = new Pen(Color.Gray);
            using var txtBrush = new SolidBrush(Color.Black);
            g.FillRectangle(back, btnRect);
            g.DrawRectangle(border, btnRect);
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString("⋯", grid.Font, txtBrush, btnRect, sf);

            e.Handled = true;
        }
        private void AbrirDiasDialog(int rowIndex)
        {
            _openingDias = true;
            try
            {
                int colIndex = grid.Columns[ColQtdDias]?.Index ?? -1;
                if (colIndex < 0) return;

                var cell = grid.Rows[rowIndex].Cells[colIndex];
                string atual = Convert.ToString(cell.Value) ?? "";
                var rect = grid.GetCellDisplayRectangle(colIndex, rowIndex, true);
                int w = Math.Max(rect.Width, 380);
                int h = 150;
                var screenPoint = grid.PointToScreen(new Point(rect.Left, rect.Bottom));

                var novoValor = DiasInputDialog.ShowAt(screenPoint, w, h,
                    "Quantidade de dias (1–999)", atual, this);

                if (novoValor == null) return;

                cell.Value = novoValor; // validado

                if (grid.IsCurrentCellInEditMode) grid.EndEdit();
                grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                grid.UpdateCellValue(colIndex, rowIndex);
                grid.InvalidateCell(colIndex, rowIndex);
                grid.Refresh();

                SetDiasEnabledFilled(rowIndex);

                AplicarAutoLarguraResponsiva();
            }
            finally { _openingDias = false; _lastDiasOpen = DateTime.UtcNow; }
        }

        // ===== Helpers visuais =====
        private bool IsPlaceholder(string? v)
        {
            var t = (v ?? "").Trim();
            return string.IsNullOrEmpty(t) || t == DISABLED_DASH;
        }

        private void SetPerfisPlaceholder(int rowIndex)
        {
            var cell = grid.Rows[rowIndex].Cells[ColPerfisAbrangentes];
            cell.Value = DISABLED_DASH;
            cell.Style.ForeColor = Color.DimGray;
            cell.Style.BackColor = Color.Gainsboro;
            cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
        private void SetPerfisNormalStyle(int rowIndex)
        {
            var cell = grid.Rows[rowIndex].Cells[ColPerfisAbrangentes];
            cell.Style.ForeColor = Color.Black;
            cell.Style.BackColor = Color.White;
            cell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
        }
        private void SetPerfisDisabled(int rowIndex)
        {
            var cell = grid.Rows[rowIndex].Cells[ColPerfisAbrangentes];
            cell.Value = DISABLED_DASH;
            cell.Style.ForeColor = Color.DimGray;
            cell.Style.BackColor = Color.Gainsboro;
            cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void SetUsuarioEnabledPlaceholder(int rowIndex)
        {
            var cell = grid.Rows[rowIndex].Cells[ColUsuarioEspecifico];
            cell.Value = DISABLED_DASH;
            cell.Style.ForeColor = Color.DimGray;
            cell.Style.BackColor = Color.White;
            cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
        private void SetUsuarioEnabledFilled(int rowIndex)
        {
            var cell = grid.Rows[rowIndex].Cells[ColUsuarioEspecifico];
            cell.Style.ForeColor = Color.Black;
            cell.Style.BackColor = Color.White;
            cell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
        }
        private void SetUsuarioDisabled(int rowIndex)
        {
            var cell = grid.Rows[rowIndex].Cells[ColUsuarioEspecifico];
            cell.Value = DISABLED_DASH;
            cell.Style.ForeColor = Color.DimGray;
            cell.Style.BackColor = Color.Gainsboro;
            cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void SetDiasEnabledPlaceholder(int rowIndex)
        {
            var cell = grid.Rows[rowIndex].Cells[ColQtdDias];
            cell.Value = "modificar";
            cell.Style.ForeColor = Color.DimGray;
            cell.Style.BackColor = Color.White;
            cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
        private void SetDiasEnabledFilled(int rowIndex)
        {
            var cell = grid.Rows[rowIndex].Cells[ColQtdDias];
            cell.Style.ForeColor = Color.Black;
            cell.Style.BackColor = Color.White;
            cell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
        }
        private void SetDiasDisabled(int rowIndex)
        {
            var cell = grid.Rows[rowIndex].Cells[ColQtdDias];
            cell.Value = DISABLED_DASH;
            cell.Style.ForeColor = Color.DimGray;
            cell.Style.BackColor = Color.Gainsboro;
            cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void AtualizarHabilitacaoLinha(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= grid.Rows.Count) return;

            var resp = Convert.ToString(grid.Rows[rowIndex].Cells[ColResponsabilidade].Value) ?? "";
            var fluxo = Convert.ToString(grid.Rows[rowIndex].Cells[ColNomeFluxo].Value) ?? "";

            // Perfis
            if (string.Equals(resp, OP_DIST, StringComparison.OrdinalIgnoreCase))
            {
                var txt = Convert.ToString(grid.Rows[rowIndex].Cells[ColPerfisAbrangentes].Value) ?? "";
                if (IsPlaceholder(txt)) SetPerfisPlaceholder(rowIndex);
                else SetPerfisNormalStyle(rowIndex);
            }
            else SetPerfisDisabled(rowIndex);

            // Usuário
            if (string.Equals(resp, OP_USER, StringComparison.OrdinalIgnoreCase))
            {
                var txt = Convert.ToString(grid.Rows[rowIndex].Cells[ColUsuarioEspecifico].Value) ?? "";
                if (IsPlaceholder(txt)) SetUsuarioEnabledPlaceholder(rowIndex);
                else SetUsuarioEnabledFilled(rowIndex);
            }
            else SetUsuarioDisabled(rowIndex);

            // Dias
            if (string.Equals(fluxo, "Encerramento", StringComparison.OrdinalIgnoreCase))
            {
                var txt = Convert.ToString(grid.Rows[rowIndex].Cells[ColQtdDias].Value) ?? "";
                if (int.TryParse(txt, out _)) SetDiasEnabledFilled(rowIndex);
                else SetDiasEnabledPlaceholder(rowIndex);
            }
            else SetDiasDisabled(rowIndex);

            int perfisCol = grid.Columns[ColPerfisAbrangentes]?.Index ?? -1;
            int userCol = grid.Columns[ColUsuarioEspecifico]?.Index ?? -1;
            int diasCol = grid.Columns[ColQtdDias]?.Index ?? -1;
            if (perfisCol >= 0) grid.InvalidateCell(perfisCol, rowIndex);
            if (userCol >= 0) grid.InvalidateCell(userCol, rowIndex);
            if (diasCol >= 0) grid.InvalidateCell(diasCol, rowIndex);
        }
        private void AtualizarHabilitacaoParaTodas()
        {
            for (int i = 0; i < grid.Rows.Count; i++)
                AtualizarHabilitacaoLinha(i);
        }

        // ======= Responsividade: medir conteúdo e preencher largura =======
        private void AplicarAutoLarguraResponsiva()
        {
            if (grid.Columns.Count == 0) return;

            grid.SuspendLayout();
            try
            {
                // mede o preferido por conteúdo e usa como MinimumWidth + FillWeight
                foreach (DataGridViewColumn c in grid.Columns)
                {
                    // mede largura ideal (conteúdo + header)
                    int pref = Math.Max(
                        c.GetPreferredWidth(DataGridViewAutoSizeColumnMode.AllCells, true), 40);

                    // colunas com botão embutido precisam de espaço extra
                    int extra = 6;
                    if (c.Name == ColPerfisAbrangentes || c.Name == ColUsuarioEspecifico || c.Name == ColQtdDias)
                        extra += (BTN_PAD + BTN_WIDTH + BTN_PAD) + TEXT_RIGHT_PAD;

                    c.MinimumWidth = pref + extra;
                    c.FillWeight = c.MinimumWidth; // proporcional ao “tamanho por caracteres”
                    c.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                }

                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            finally { grid.ResumeLayout(); }
        }

        // ================= Botões inferiores =================
        private void LimparPreenchimento()
        {
            foreach (DataGridViewRow row in grid.Rows)
            {
                row.Cells[ColAtivacao].Value = "Selecione";
                row.Cells[ColDetalharRegra].Value = "Não";
                row.Cells[ColResponsabilidade].Value = "Selecione";
                row.Cells[ColPerfisAbrangentes].Value = "";
                row.Cells[ColUsuarioEspecifico].Value = "";
                row.Cells[ColQtdDias].Value = "";
            }
            AplicarOpcoesResponsabilidadePorFluxo();
            AtualizarHabilitacaoParaTodas();
            grid.Refresh();
        }

        private void GerarExcel()
        {
            MessageBox.Show(this,
                "Aqui você aciona o seu gerador de Excel.\n(Troque o método GerarExcel() pelo seu código atual.)",
                "Gerar Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static GraphicsPath RoundedRect(Rectangle bounds) => RoundedRect(bounds, 8);

        private void MostrarColuna(string colName, bool visivel)
        {
            if (grid.Columns.Contains(colName))
                grid.Columns[colName].Visible = visivel;
        }
        private void LiberarColunaAtivacao() => MostrarColuna(ColAtivacao, true);
        private void LiberarColunaDetalharRegra() => MostrarColuna(ColDetalharRegra, true);
        private void LiberarColunaResponsabilidade() => MostrarColuna(ColResponsabilidade, true);
    }

    // ========= Dropdown MULTI-SELECT com busca (Perfis) =========
    internal sealed class PerfisDropDown : Form
    {
        private readonly Panel _header;
        private readonly Label _headerText;
        private readonly TextBox _search;
        private readonly Button _clearSearch;
        private readonly CheckedListBox _list;
        private readonly Button _ok;
        private readonly Button _cancel;

        private readonly string[] _allOptions;
        private readonly HashSet<string> _checked;

        private PerfisDropDown(string[] opcoes, IEnumerable<string> preSelecionados)
        {
            _allOptions = opcoes ?? Array.Empty<string>();
            _checked = new HashSet<string>(preSelecionados ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);

            Text = "";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            ShowInTaskbar = false;
            MinimizeBox = false; MaximizeBox = false;
            TopMost = true; KeyPreview = true;

            _header = new Panel { Left = 0, Top = 0, Width = 450, Height = 36, BackColor = Color.FromArgb(245, 245, 245) };
            _headerText = new Label { AutoSize = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold) };
            _header.Controls.Add(_headerText);

            _search = new TextBox { PlaceholderText = "Pesquisar...", Left = 8, Top = _header.Bottom + 6, Width = 360 };
            _clearSearch = new Button { Text = "✕", Left = _search.Right + 6, Top = _header.Bottom + 5, Width = 28, Height = 26 };
            _clearSearch.Click += (s, e) => { _search.Text = ""; _search.Focus(); };

            _list = new CheckedListBox { Left = 8, Top = _search.Bottom + 6, Width = 420, Height = 230, CheckOnClick = true, IntegralHeight = false };
            _list.ItemCheck += (s, e) =>
            {
                var itemText = _list.Items[e.Index]?.ToString() ?? "";
                void Update()
                {
                    if (e.NewValue == CheckState.Checked) _checked.Add(itemText);
                    else _checked.Remove(itemText);
                    UpdateHeader();
                }
                if (IsHandleCreated) BeginInvoke((Action)Update);
                else Update();
            };

            _ok = new Button { Text = "OK", Width = 100, Height = 30, Left = 228, Top = _list.Bottom + 8, DialogResult = DialogResult.OK };
            _cancel = new Button { Text = "Cancelar", Width = 100, Height = 30, Left = 336, Top = _list.Bottom + 8, DialogResult = DialogResult.Cancel };
            AcceptButton = _ok; CancelButton = _cancel;

            Controls.AddRange(new Control[] { _header, _search, _clearSearch, _list, _ok, _cancel });

            _search.TextChanged += (s, e) =>
            {
                ApplyFilter(_search.Text);
                for (int i = 0; i < _list.Items.Count; i++)
                {
                    var txt = _list.Items[i]?.ToString() ?? "";
                    _list.SetItemChecked(i, _checked.Contains(txt));
                }
            };

            ApplyFilter("");
            for (int i = 0; i < _list.Items.Count; i++)
            {
                var txt = _list.Items[i]?.ToString() ?? "";
                if (_checked.Contains(txt)) _list.SetItemChecked(i, true);
            }
            UpdateHeader();

            KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); }
                if (e.KeyCode == Keys.Enter && !_search.Focused) { DialogResult = DialogResult.OK; Close(); }
            };

            LayoutControls();
            Resize += (s, e) => LayoutControls();
        }

        private void LayoutControls()
        {
            _header.Width = ClientSize.Width;
            _search.Width = ClientSize.Width - 8 - 8 - _clearSearch.Width - 6;
            _clearSearch.Left = _search.Right + 6;
            _list.Width = ClientSize.Width - 16;
            _list.Height = ClientSize.Height - _list.Top - 46;
            _cancel.Top = _ok.Top = _list.Bottom + 8;
            _cancel.Left = ClientSize.Width - 8 - _cancel.Width;
            _ok.Left = _cancel.Left - 8 - _ok.Width;
        }

        private void ApplyFilter(string term)
        {
            term = (term ?? "").Trim();
            _list.BeginUpdate();
            _list.Items.Clear();

            IEnumerable<string> data = _allOptions;
            if (!string.IsNullOrEmpty(term))
                data = data.Where(x => x?.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0);

            foreach (var it in data) _list.Items.Add(it, _checked.Contains(it));
            _list.EndUpdate();
        }

        private void UpdateHeader()
        {
            int n = _checked.Count;
            _headerText.Text = n switch
            {
                0 => "Nenhum item selecionado ▾",
                1 => "1 item selecionado ▾",
                _ => $"{n} itens selecionados ▾"
            };
        }

        private string[] GetSelected() => _allOptions.Where(o => _checked.Contains(o)).ToArray();

        public static string[]? ShowAt(Point screenPoint, int width, int height,
                                       string[] opcoes, IEnumerable<string> preSel, IWin32Window owner)
        {
            using var dd = new PerfisDropDown(opcoes, preSel) { StartPosition = FormStartPosition.Manual };
            dd.Width = Math.Max(380, width);
            dd.Height = Math.Max(320, height);

            var wa = Screen.FromPoint(screenPoint).WorkingArea;
            int x = screenPoint.X;
            int y = screenPoint.Y;
            if (y + dd.Height > wa.Bottom) y = screenPoint.Y - dd.Height - 4;
            if (x < wa.Left) x = wa.Left + 4;
            if (x + dd.Width > wa.Right) x = wa.Right - dd.Width - 4;
            if (y < wa.Top) y = wa.Top + 4;

            dd.Location = new Point(x, y);

            var result = dd.ShowDialog(owner);
            if (result != DialogResult.OK) return null;
            return dd.GetSelected();
        }
    }

    // ========= Diálogo de texto (Usuário Específico) =========
    internal sealed class UsuarioInputDialog : Form
    {
        private readonly Label _title;
        internal readonly TextBox _txt;
        private readonly Button _ok;
        private readonly Button _cancel;

        private UsuarioInputDialog(string titulo, string valorInicial)
        {
            Text = "";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            ShowInTaskbar = false;
            MinimizeBox = false; MaximizeBox = false;
            TopMost = true; KeyPreview = true;

            _title = new Label { Text = titulo, AutoSize = false, Dock = DockStyle.Top, Height = 28, TextAlign = ContentAlignment.MiddleCenter, Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold) };
            _txt = new TextBox { Left = 10, Top = _title.Bottom + 8, Width = 320, Text = valorInicial ?? "" };
            _ok = new Button { Text = "OK", Width = 100, Height = 30, Left = 140, Top = _txt.Bottom + 10, DialogResult = DialogResult.OK };
            _cancel = new Button { Text = "Cancelar", Width = 100, Height = 30, Left = _ok.Right + 8, Top = _txt.Bottom + 10, DialogResult = DialogResult.Cancel };

            AcceptButton = _ok;
            CancelButton = _cancel;

            Controls.AddRange(new Control[] { _title, _txt, _ok, _cancel });

            Width = 360;
            Height = _cancel.Bottom + 50;

            _txt.KeyPress += (s, e) =>
            {
                if (e.KeyChar == '\r' || e.KeyChar == '\n')
                {
                    DialogResult = DialogResult.OK;
                    Close();
                }
            };
        }

        public static string? ShowAt(Point screenPoint, int width, int height, string titulo, string valorInicial, IWin32Window owner)
        {
            using var dlg = new UsuarioInputDialog(titulo, valorInicial) { StartPosition = FormStartPosition.Manual };
            dlg.Location = screenPoint;
            dlg.Width = Math.Max(320, width);
            dlg.Height = Math.Max(120, height);

            var result = dlg.ShowDialog(owner);
            if (result != DialogResult.OK) return null;
            return (dlg._txt.Text ?? "").Trim();
        }
    }

    // ========= Diálogo numérico (1..999) para Dias =========
    internal sealed class DiasInputDialog : Form
    {
        private readonly Label _title;
        private readonly TextBox _txt;
        private readonly Button _ok;
        private readonly Button _cancel;

        private DiasInputDialog(string titulo, string valorInicial)
        {
            Text = "";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            ShowInTaskbar = false;
            MinimizeBox = false; MaximizeBox = false;
            TopMost = true; KeyPreview = true;

            _title = new Label { Text = titulo, AutoSize = false, Dock = DockStyle.Top, Height = 28, TextAlign = ContentAlignment.MiddleCenter, Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold) };
            _txt = new TextBox { Left = 10, Top = _title.Bottom + 8, Width = 200, Text = valorInicial?.Trim().ToLower() == "modificar" ? "" : (valorInicial ?? "") };
            _ok = new Button { Text = "OK", Width = 90, Height = 28, Left = 120, Top = _txt.Bottom + 10, DialogResult = DialogResult.OK };
            _cancel = new Button { Text = "Cancelar", Width = 90, Height = 28, Left = _ok.Right + 8, Top = _txt.Bottom + 10, DialogResult = DialogResult.Cancel };

            AcceptButton = _ok;
            CancelButton = _cancel;

            Controls.AddRange(new Control[] { _title, _txt, _ok, _cancel });

            Width = 300;
            Height = _cancel.Bottom + 50;

            _txt.MaxLength = 3;
            _txt.KeyPress += (s, e) =>
            {
                if (char.IsControl(e.KeyChar)) return;
                if (!char.IsDigit(e.KeyChar)) e.Handled = true;
            };

            _ok.Click += (s, e) =>
            {
                if (!int.TryParse(_txt.Text.Trim(), out var n) || n < 1 || n > 999)
                {
                    MessageBox.Show(this, "Informe um número entre 1 e 999.", "Valor inválido",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.DialogResult = DialogResult.None;
                    _txt.Focus();
                    _txt.SelectAll();
                }
            };
        }

        public static string? ShowAt(Point screenPoint, int width, int height, string titulo, string valorInicial, IWin32Window owner)
        {
            using var dlg = new DiasInputDialog(titulo, valorInicial) { StartPosition = FormStartPosition.Manual };
            dlg.Location = screenPoint;
            dlg.Width = Math.Max(280, width);
            dlg.Height = Math.Max(130, height);

            var result = dlg.ShowDialog(owner);
            if (result != DialogResult.OK) return null;
            return dlg._txt.Text.Trim();
        }
    }
}
