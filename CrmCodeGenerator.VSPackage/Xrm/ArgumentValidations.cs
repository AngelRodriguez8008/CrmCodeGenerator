using System;

namespace CrmCodeGenerator.VSPackage.Xrm
{
	public static class  ArgumentValidations
	{
		//[ContractAnnotation("entityId:null => false")]
		public static bool IsValid(this Guid? entityId) => entityId.IsNullOrEmpty() == false;
		
		//[ContractAnnotation("entityId:null => true")]
		public static bool IsNullOrEmpty(this Guid? entityId)
		{
			return entityId == null || entityId == default(Guid);
		}
		
		public static bool IsValid(this string source)
		{
			return !string.IsNullOrWhiteSpace(source);
		}

		public static bool IsValid(this Enum source)
		{
			return source != null && Enum.IsDefined(source.GetType(), source);
		}
    
		public static Guid Valid(this Guid? source, string parameter, string errorMessage = null)
		{
			if (errorMessage == null)
				errorMessage = "This guid id can not be null nor empty";

			return source.IsValid() ? source.Value : throw new ArgumentNullException(parameter, errorMessage);
		}
		
		public static string Valid(this string source, string parameter, string errorMessage = null)
		{
			if (errorMessage == null)
				errorMessage = "This string value can not be null nor empty";

			return IsValid(source) ? source : throw new ArgumentNullException(parameter, errorMessage);
		}

		public static T Valid<T>(this T source, string parameter, string errorMessage = null)  where T : class
		{
			if (errorMessage == null)
				errorMessage = "This value can not be null";
			return source ?? throw new ArgumentNullException(parameter, errorMessage);
		}

		public static T Valid<T>(this T source) where T : class => source.Valid(null);
		

		public static T Valid<T, TException>(this T source)
			where T : class
			where TException : Exception, new()
		{
			return source ?? throw new TException();
		}
	}
}