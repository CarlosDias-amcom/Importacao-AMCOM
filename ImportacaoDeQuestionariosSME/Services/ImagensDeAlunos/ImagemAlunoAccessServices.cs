using ImportacaoDeQuestionariosSME.Data.Repositories.ImagensAlunos;
using ImportacaoDeQuestionariosSME.Domain.ImagensAlunos;
using ImportacaoDeQuestionariosSME.Services.ImagensDeAlunos.Dtos;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportacaoDeQuestionariosSME.Services.ImagensDeAlunos
{
    public class ImagemAlunoAccessServices : IImagemAlunoServices
    {
        private readonly IImagemAlunoRepository _imagemAlunoRepository;

        public ImagemAlunoAccessServices()
        {
            _imagemAlunoRepository = new ImagemAlunoRepository();
        }

        public async Task ImportarAsync(ImportacaoDeImagemAlunoDto dto)
        {
            if (dto is null)
            {
                dto = new ImportacaoDeImagemAlunoDto();
                dto.AddErro("O DTO é nulo.");
                return;
            }

            try
            {
                var connection = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dto.CaminhoDaDoArquivo}";
                var dtImagens = GetImagensAlunoFromAccess(connection);
                if (dtImagens.Rows.Count <= 0)
                {
                    dto.AddErro("Não existem regitros na planilha para exportação.");
                    return;
                }

                var entities = dtImagens
                    .AsEnumerable()
                    .Select(row => new ImagemAluno
                    {
                        AluMatricula = row["CD_ALUNO_SME"].ToString(),
                        AluNome = row["NOME"].ToString(),
                        AreaConhecimentoId = GetAreaDeConhecimentoId(row["DISCIPLINA"].ToString()),
                        Caminho = FormatCaminhoDaImagem(row["IMAGEM"].ToString()),
                        Edicao = dto.Ano,
                        EscCodigo = row["CD_UNIDADE_EDUCACAO"].ToString(),
                        Pagina = int.Parse(row["PAGINA"].ToString()),
                        Questao = row["QUESTAO"].ToString()
                    })
                    .ToList();

                AdjustEntities(entities);

                await _imagemAlunoRepository.InsertAsync(entities);
            }
            catch (Exception ex)
            {
                dto.AddErro(ex.InnerException?.Message ?? ex.Message);
            }
        }

        private string GetQuestao(string disciplina)
        {
            switch (disciplina)
            {
                case "MT":
                    return "Matemática";
                case "LP":
                    return "Língua Portuguesa";
                case "CI":
                    return "Ciências da Natureza";
                default:
                    return "Redação";
            }
        }

        private DataTable GetImagensAlunoFromAccess(string connString)
        {
            var query = @"SELECT CadernosRed.INSC, CadernosRed.CD_ALUNO_SME, CadernosRed.NOME, CadernosRed.CD_UNIDADE_EDUCACAO, CadernosRed.QUESTAO, CadernosRed.PAGINA, CadernosRed.DISCIPLINA, CadernosRed.IMAGEM
                        FROM CadernosRed
                        WHERE CadernosRed.CD_ALUNO_SME <>''
                        AND CadernosRed.QUESTAO <> ''
                        AND CadernosRed.DISCIPLINA = 'MT'";

            var dAdapter = new OleDbDataAdapter(query, connString);
            var dTable = new DataTable();
            var cBuilder = new OleDbCommandBuilder(dAdapter);
            cBuilder.QuotePrefix = "[";
            cBuilder.QuoteSuffix = "]";
            dAdapter.Fill(dTable);
            return dTable;
        }

        private static int GetAreaDeConhecimentoId(string disciplina)
        {
            switch (disciplina)
            {
                case "MT":
                    return 3;
                case "LP":
                    return 2;
                case "CI":
                    return 1;
                default:
                    return 4;
            }
        }

        private string FormatCaminhoDaImagem(string caminhoDaImagem)
        {
            caminhoDaImagem = caminhoDaImagem.Replace("\\", "/");

            var indice = caminhoDaImagem.IndexOf("IMAGEM/") + 7;
            caminhoDaImagem = $"IMAGENS/2019/{caminhoDaImagem.Substring(indice, caminhoDaImagem.Length - indice)}";

            return caminhoDaImagem;
        }

        private void AdjustEntities(IList<ImagemAluno> entities)
        {
            var duplicates = entities
                .GroupBy(x => new { x.AluMatricula, x.Questao, x.Pagina })
                .Where(x => x.Count() > 1)
                .ToList();

            if (!duplicates.Any()) return;

            foreach (var grouping in duplicates)
                foreach (var entity in grouping)
                    entities.Remove(entity);
        }
    }
}
