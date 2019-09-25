using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace Exemplo
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Criarcao do documento

                //Cria um documento com nome Exemplodoc!
                Document ExemploDoc = new Document();
            #endregion

            #region
                //adiciona uma seção com nome secacAPA ao documento
                //cada secao pode ser entendida como uma pagina do documento 
                Section secaoCapa = ExemploDoc.AddSection();
            #endregion

            #region Criar um paragrafo
                //Cria um paragrafo com o nome titulo e adicioa à seção secaoCapa
                //Os paragrafos são necessários para inserção de texto, imagens, etc
                Paragraph titulo = secaoCapa.AddParagraph();               
            #endregion
            
            #region Adiciona texto ao paragrafo
                //adiciona o texto exemplom de titulo ao paragrafo titulo
                titulo.AppendText("Exemplo de título\n\n");
            #endregion 

            #region Formatar paragrafo
                //Através da propriedade HorizontalAlignment, é possível alinhar o parágrafo
                titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

                ParagraphStyle estilo01 = new ParagraphStyle (ExemploDoc);

                //adiciona um nome ao estilo
                estilo01.Name = "Cor do titulo";

                //defini a cor do titulo
                estilo01.CharacterFormat.TextColor = Color.Pink;

                //define que o texto seá em negrito
                estilo01.CharacterFormat.Bold = true;

                //Adiciona estilo01 ao exemploDoc
                ExemploDoc.Styles.Add(estilo01);

                //aplica o estilo 01 ao parágrafo titulo
                titulo.ApplyStyle(estilo01.Name);
            #endregion

            #region Trabalhar taulação 
            //Adiciona um paragrafo textoCapa à seção  secaoCapa
            Paragraph textoCapa = secaoCapa.AddParagraph();
                
             //adiciona um texto no paragrafo com tabulação   
            textoCapa.AppendText("\tEste é um exemplo de texto com tabulaçãoz\n");

            //adiciona um novo parágrafo á mesma seção (secaoCapa)
            Paragraph textocapa2 = secaoCapa.AddParagraph();

            //adiciona um texto ao parágrafo textoCapa2 com concatenação
            textocapa2.AppendText("\tBasicamente, então, uma selão representa uma pagina do documento," + "oobviamente, aparecem na mesma página");
            #endregion

            #region Inserir imagens
            //Adiciona um parágrafo à seção secaoCapa
            Paragraph ImagemCapa = secaoCapa.AddParagraph();

            //Adiciona um texto ao parágrafo ImagemCapa    
            ImagemCapa.AppendText("\n\n\tAgora vamos insirir uma imagem ao documento\n\n");
            
            //centraliza horizontalmente o parágrafo imagemCapa
            ImagemCapa.Format.HorizontalAlignment = HorizontalAlignment.Center;

            //Adiciona uma imagem com o nome imagemExemplo ao parágrafo Imagemcapa
            DocPicture imagemExemplo = ImagemCapa.AppendPicture(Image.FromFile(@"saida\img\logo_csharp.png"));
            #endregion
            

            //Define uma largura e uma altura para imagem
            imagemExemplo.Width = 300;
            imagemExemplo.Height = 300;

            #region  Adicionar nova seção
            //Adiciona uma nova seção 
            Section SecaoCorpo = ExemploDoc.AddSection();

            //Adiciona um parágrafo à seção secaocorpo
            Paragraph paragrafoCorpo1 = SecaoCorpo.AddParagraph();

            //Adiciona um texto ao paragrafoCorpo1
            paragrafoCorpo1.AppendText("\t este é um exemplo de parágrafo criado em uma nova seção" + "\t Como foi criada uma nova seção, perceba que este texto aparece em uma nova página");
            #endregion

            #region Adicionar uma tabela
            //Adiciona uma tabela à seção secaoCorpo
            Table tabela = SecaoCorpo.AddTable(true);

            //Cria o cabeçalho da tabela
            String[] cabeçalho = {"Item", "Descrição", "Qtd", "Preço Unit", "Preço"};

            String [] [] dados = {
                new String []{"Cenoura","Vegetal muito nutritivo","1","R$ 4,00","R$ 4,00"},
                new String []{"Batata","Vegetal muito Consumido","2","R$ 5,00","R$ 10,00"},
                new String []{"Alface","Vegetal muito utilizade desde 500 a.C.","1","R$ 1,50","R$ 1,50"},
                new String []{"Tomate","Tomate é uma é uma fruta","2","R$ 6,00","R$ 12,00"}
            };

            //Adiciona as células na tabela
            tabela.ResetCells(dados.Length + 1, cabeçalho.Length);
            
            //Adiciona uma linha na posição 0 do vertor de linha
            TableRow Linha1 = tabela.Rows[0]; 
            Linha1.IsHeader = true;

            //Define a altura da linha
            Linha1.Height = 23;

            //Formatação do cabeçalho
            Linha1.RowFormat.BackColor = Color.AliceBlue;

            //percorre as colunas do cabeçalho
            for (int i = 0; i < cabeçalho.Length; i++)
            {
                //Alinhamneto das células
                Paragraph p = Linha1.Cells[i].AddParagraph();
                Linha1.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                //Formatação dos dados do cabeçalho
                TextRange TR = p.AppendText(cabeçalho[i]);
                TR.CharacterFormat.FontName = "Calibri";
                TR.CharacterFormat.FontSize = 14;
                TR.CharacterFormat.TextColor = Color.Teal;
                TR.CharacterFormat.Bold = true;

            }

            //Adiciona as linhas do corpo da tabela
            for (int r = 0; r < dados.Length; r++)
            {
                TableRow LinhaDados = tabela.Rows[r + 1];

                //Define altura da linha
                LinhaDados.Height = 20;

                //percorre as Colunas
                for (int c = 0; c < dados[r].Length; c++)
                {
                    //Alinha as células
                    LinhaDados.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                    //preenche os dados nas linhas
                    Paragraph p2 = LinhaDados.Cells[c].AddParagraph();
                    TextRange TR2 = p2.AppendText(dados[r] [c]);

                    //Formata as células
                    p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR2.CharacterFormat.FontName = "Calibri";
                    TR2.CharacterFormat.FontSize = 12;
                    TR2.CharacterFormat.TextColor = Color.Brown;
                }
                
            }
            
            #endregion

            #region Salvar arquivo
                //Salvar o arquivo em .Docx
                //utiliza o método SaveToFile para salvar o arquivo no formato desejado
                //Assim como no word, caso já existia um arquivo com este nome, é substituído
                ExemploDoc.SaveToFile(@"saida\exemplo_arquivo_word.docx", FileFormat.Docx); 
                
            #endregion
        }
    }
}
