using System;

namespace Exercicios
{
    class Program
    {
        static void Main(string[] args)
        {
            #region criacao de documento
                //Cria um documento com nome Exemplodoc
                Document ExemploDoc = new Document();

            #endregion

            #region Cria uma seção
                Section secaoCapa = ExemploDoc.addSection();   
            #endregion
            

        }
    }
}
