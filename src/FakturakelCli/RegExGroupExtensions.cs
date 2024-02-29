using System.Text.RegularExpressions;

public static class RegExGroupExtensions
{
    public static decimal ToDecimal(this Group group)
    {
        return Convert.ToDecimal(group.Value);
    }
}