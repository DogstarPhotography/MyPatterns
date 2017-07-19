// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Data;
using Microsoft.VisualBasic;
using System.Collections;
// End of VB project level imports


/// <summary>
///
/// </summary>
/// <typeparam name="T"></typeparam>
/// <remarks>See http://stackoverflow.com/questions/940002/getting-static-field-values-of-a-type-using-reflection/940340#940340</remarks>
namespace Generics
{
	public interface ICustomEnum<T>
	{
		ICustomEnum<T> FromT(T value);
		T Value {get;}
		//'// Implement using a private constructor that accepts and sets the Value property,
		//'// one shared readonly property for each desired value in the enum,
		//'// and widening conversions to and from T.
		//'// Then see this link to get intellisense support
		//'// that exactly matches a normal enum:
		//'// http://stackoverflow.com/questions/102084/hidden-features-of-vb-net/102217#102217
	}
	
	/// <completionlist cref="ReasonCodeValue"/>
	public sealed class ReasonCodeValue : ICustomEnum<string>
	{
		
		private string _value;
public string Value
		{
			get
			{
				return _value;
			}
		}
		
		private ReasonCodeValue(string value)
		{
			_value = value;
		}
		
		private static ReasonCodeValue _ServiceNotCovered = new ReasonCodeValue("SNCV");
public static ReasonCodeValue ServiceNotCovered
		{
			get
			{
				return _ServiceNotCovered;
			}
		}
		
		private static ReasonCodeValue _MemberNotEligible = new ReasonCodeValue("MNEL");
public static ReasonCodeValue MemberNotEligible
		{
			get
			{
				return _MemberNotEligible;
			}
		}
		
		public static ICustomEnum<string> FromString(string value)
		{
			//'// use reflection or a dictionary here if you have a lot of values
			switch (value)
			{
				case "SNCV":
					return _ServiceNotCovered;
				case "MNEL":
					return _MemberNotEligible;
				default:
					return default(ICustomEnum<string>); //'//or throw an exception
			}
		}
		
		public ICustomEnum<string> FromT(string value)
		{
			return FromString(value);
		}
		
		public static implicit operator string (ReasonCodeValue item)
		{
			return item.Value;
		}
		
		public static implicit operator ReasonCodeValue (string value)
		{
			return ((ReasonCodeValue) (FromString(value)));
		}
	}
}
