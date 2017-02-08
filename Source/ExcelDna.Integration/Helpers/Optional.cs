using System;

namespace ExcelDna.Integration.Helpers
{

  public static class Optional
  {
    public static T Value<T>(object arg, T defaultValue, string paramName) {
      if (arg is T) {
        return (T)arg;
      }
      else if (arg is ExcelMissing) {
        return defaultValue;
      }
      else
        throw new ArgumentException(String.Concat(
          "Invalid optional argument type: ", paramName, " should be a ", defaultValue.GetType().FullName,
          " but is a ", arg.GetType().FullName, "."
        ));
    }
  }

}
