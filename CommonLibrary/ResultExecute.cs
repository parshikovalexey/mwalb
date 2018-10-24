using HelperLibrary;
using System;
using System.ComponentModel.DataAnnotations;

namespace CommonLibrary
{
    public class ResultExecute
    {
        public enum Status
        {
            [Display(Name = "Ошибка")]
            Error = 0,
            [Display(Name = "Успех")]
            Success = 1,
            [Display(Name = "Внимание")]
            Warm = 2
        }
        public string ErrorMsg { get; set; } = String.Empty;
        public Status StatusExecute { get; set; } = Status.Success;
        public object Callbacks { get; set; }

        public override string ToString()
        {
            return $"StatusExecute: {StatusExecute}, ErrorMsg: {ErrorMsg}, Obj {Callbacks}";
        }

        public ResultExecute OnError(string msg, object obj = null)
        {
            return new ResultExecute() { Callbacks = obj, ErrorMsg = msg, StatusExecute = Status.Error };
        }

        public ResultExecute OnSuccess(object obj = null)
        {
            return new ResultExecute() { Callbacks = obj, StatusExecute = Status.Success };
        }

        public ResultExecute OnSuccess(string msg)
        {
            return new ResultExecute() { Callbacks = null, ErrorMsg = msg, StatusExecute = Status.Success };
        }

        public string GetMsg() => string.IsNullOrEmpty(ErrorMsg) ? $"{StatusExecute.DisplayName()}" : $"{StatusExecute.DisplayName()} {ErrorMsg}";
    }
}
