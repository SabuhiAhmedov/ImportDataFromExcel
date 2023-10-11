namespace DataReadFromExcel.Helper
{
    public static class Extensions
    {
        public static int GetKeyByValue(this Dictionary<int,string> dictionary,string value)
        {
            int myKey = dictionary.FirstOrDefault(x => x.Value == value).Key;
            return myKey;
        }
    }
}
