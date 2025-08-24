using System;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace PerformaTools
{
    internal static class ExportadorExcel
    {
        // nomes das colunas no grid (devem ser iguais aos usados no form)
        private const string ColNomeFluxo = "colNomeFluxo";
        private const string ColAtivacao = "colAtivacao";
        private const string ColDetalharRegra = "colDetalharRegra";
        private const string ColResponsabilidade = "colResponsabilidade";
        private const string ColPerfisAbrangentes = "colPerfisAbrangentes";
        private const string ColUsuarioEspecifico = "colUsuarioEspecifico";
        private const string ColQtdDias = "colQtdDias";

        // valores “padrão” da UI (sem alteração)
        private const string PADRAO_ATIVACAO = "Selecione";
        private const string PADRAO_RESP = "Selecione";
        private const string PADRAO_DETALHAR = "Não";    // não é editável
        private const string PLACEHOLDER = "—";      // placeholder em campos desabilitados/sem valor
        private const string PADRAO_DIAS = "modificar"; // placeholder para Encerramento

        // estilos Excel
        private static readonly XLColor HeaderBg = XLColor.FromArgb(0, 32, 96);
        private static readonly XLColor HeaderFg = XLColor.White;
        private static readonly XLColor ChangedFill = XLColor.FromArgb(255, 242, 204); // amarelo suave #FFF2CC

        public static void ExportRegraGeral(DataGridView grid, IWin32Window owner)
        {
            if (grid == null || grid.Columns.Count == 0)
            {
                MessageBox.Show(owner, "Não há dados para exportar.", "Gerar Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Escolha do arquivo
            using var sfd = new SaveFileDialog
            {
                Title = "Salvar relatório de alterações",
                Filter = "Arquivo Excel (*.xlsx)|*.xlsx",
                FileName = $"Relatorio_Alteracao_Regras_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            };
            if (sfd.ShowDialog(owner) != DialogResult.OK) return;

            // Planilha
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Alterações");

            // Cabeçalhos (usa as colunas VISÍVEIS e na ordem que aparecem no grid)
            var cols = grid.Columns
                           .Cast<DataGridViewColumn>()
                           .Where(c => c.Visible)
                           .OrderBy(c => c.DisplayIndex)
                           .ToList();

            for (int c = 0; c < cols.Count; c++)
            {
                var cell = ws.Cell(1, c + 1);
                cell.Value = cols[c].HeaderText;
                cell.Style.Fill.BackgroundColor = HeaderBg;
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontColor = HeaderFg;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            int rowExcel = 1;

            // Apenas linhas com alterações
            foreach (DataGridViewRow r in grid.Rows)
            {
                if (r.IsNewRow) continue;
                if (!LinhaTemAlteracao(r)) continue;

                rowExcel++;
                for (int c = 0; c < cols.Count; c++)
                {
                    var colName = cols[c].Name;
                    var valor = r.Cells[colName].Value;
                    var xCell = ws.Cell(rowExcel, c + 1);

                    // Escreve valor (numérico para Dias, se for o caso)
                    if (colName == ColQtdDias && valor != null && int.TryParse(Convert.ToString(valor), out var dias))
                        xCell.Value = dias; // numérico
                    else
                        xCell.Value = Convert.ToString(valor) ?? "";

                    // bordas padrão
                    xCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    // pinta de amarelo o que foi modificado
                    if (CelulaAlterada(r, colName))
                        xCell.Style.Fill.BackgroundColor = ChangedFill;
                }
            }

            if (rowExcel == 1)
            {
                MessageBox.Show(owner, "Nenhuma alteração encontrada para exportar.", "Gerar Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ws.SheetView.FreezeRows(1);
            ws.Columns().AdjustToContents(2.0, 60.0); // largura confortável
            wb.SaveAs(sfd.FileName);

            MessageBox.Show(owner, "Excel gerado com sucesso!", "Gerar Excel",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ----- regras de “alteração” -----
        private static bool LinhaTemAlteracao(DataGridViewRow r)
        {
            return CelulaAlterada(r, ColAtivacao)
                || CelulaAlterada(r, ColResponsabilidade)
                || CelulaAlterada(r, ColPerfisAbrangentes)
                || CelulaAlterada(r, ColUsuarioEspecifico)
                || CelulaAlterada(r, ColQtdDias);
        }

        private static bool CelulaAlterada(DataGridViewRow r, string colName)
        {
            string v = (Convert.ToString(r.Cells[colName].Value) ?? "").Trim();

            return colName switch
            {
                ColAtivacao => !v.Equals(PADRAO_ATIVACAO, StringComparison.OrdinalIgnoreCase),
                ColResponsabilidade => !v.Equals(PADRAO_RESP, StringComparison.OrdinalIgnoreCase),
                ColPerfisAbrangentes => !string.IsNullOrEmpty(v) && !v.Equals(PLACEHOLDER, StringComparison.Ordinal),
                ColUsuarioEspecifico => !string.IsNullOrEmpty(v) && !v.Equals(PLACEHOLDER, StringComparison.Ordinal),
                ColQtdDias => int.TryParse(v, out var n) && n >= 1 && n <= 999,
                _ => false
            };
        }
    }
}
