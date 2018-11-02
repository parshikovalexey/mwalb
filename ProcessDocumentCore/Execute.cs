using System.IO;
using System.Linq;
using CommonLibrary;
using ProcessDocumentCore.Interface;
using StandardsLibrary;

namespace ProcessDocumentCore
{

    public class Execute
    {
        public delegate void AddPreparedDocumentInterface(ResultExecute resultExecute);
        protected AddPreparedDocumentInterface PreparedDocument;

        /// <summary>
        /// Запускаем обработку документа
        /// </summary>
        /// <param name="resultDocumentInterface">Интерфейс возвращает результат обработки</param>
        /// <param name="designStandard">стандарт ГОСТ</param>
        /// <param name="filePath">Путь до файла для обработки</param>
        public Execute(string filePath, Standards designStandard, IDocumentProcessing documentProcessing, AddPreparedDocumentInterface resultDocumentInterface )
        {
            this.PreparedDocument = resultDocumentInterface;
            if (string.IsNullOrEmpty(filePath)) OnResponsePreparedDocument(new ResultExecute().OnError("Путь до файла не указан"));
            if (!PathHelper.FileExists(filePath, false)) OnResponsePreparedDocument(new ResultExecute().OnError($"Указанный файл '{filePath}' не существует"));
            if (!IsValidFile(filePath)) OnResponsePreparedDocument(new ResultExecute().OnError($"Загружен файл с неизвестным форматом '{filePath}'. Поддерживаются только файлы Word *.doc или *.docx"));
            
            //передаем данные на форматирование
            OnResponsePreparedDocument(documentProcessing.Processing(designStandard, filePath));

           //OnResponsePreparedDocument(new ResultExecute(){Callbacks = filePath});
        }

        private bool IsValidFile(string filePath)
        {
            string[] fileExtension = { ".doc", ".docx" };
            FileInfo f = new FileInfo(filePath);
            return fileExtension.Contains(f.Extension);
        }

        /// <summary>
        /// Возвращаем результат обработки документа
        /// </summary>
        /// <param name="resultExecute">Результат обработки документа</param>
        protected virtual void OnResponsePreparedDocument(ResultExecute resultExecute)
        {
            PreparedDocument?.Invoke(resultExecute);
        }
    }
}
