// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Data;
using Microsoft.VisualBasic;
using System.Collections;
// End of VB project level imports


/// <summary>
/// Can be used to return a result from a procedure where several bits of data need to be tied together,
/// such as wether the procedure failed or succeeded, what the result code was, any message, and any exception details.
/// </summary>
/// <remarks></remarks>
namespace Generics
{
	public class ProcedureResult
	{
		
		public bool Result;
		public int Code;
		public string Message;
		public Exception Exception;
		
		public ProcedureResult(bool Result, int Code, string Message, Exception Exception)
		{
			this.Result = Result;
			this.Code = Code;
			this.Message = Message;
			this.Exception = Exception;
		}
	}
	/// <summary>
	/// ProcedureResult that includes an object to allow for passing of untyped data out of a procedure
	/// </summary>
	/// <remarks></remarks>
	public partial class FunctionResult : ProcedureResult
	{
		
		public object Data;
		
		public FunctionResult(bool Result, int Code, string Message, Exception Exception, object Data) : base(Result, Code, Message, Exception)
		{
			this.Data = Data;
		}
	}
	/// <summary>
	/// ProcedureResult that includes a generic to allow for passing of typed data out of a procedure
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <remarks></remarks>
	public partial class FunctionResult<T> : ProcedureResult
	{
		
		public T Data;
		
		public FunctionResult(bool Result, int Code, string Message, Exception Exception, T Data) : base(Result, Code, Message, Exception)
		{
			this.Data = Data;
		}
	}
	
}
