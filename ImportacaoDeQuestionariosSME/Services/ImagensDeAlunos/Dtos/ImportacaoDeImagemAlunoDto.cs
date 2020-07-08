using ImportacaoDeQuestionariosSME.Services.Abstractions;

namespace ImportacaoDeQuestionariosSME.Services.ImagensDeAlunos.Dtos
{
    public class ImportacaoDeImagemAlunoDto : BaseDto
    {
        public ImportacaoDeImagemAlunoDto()
            : base()
        {
        }

        public string CaminhoDaDoArquivo { get; set; }
        public string Ano { get; set; }
    }
}