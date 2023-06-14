namespace DNSCAD
{
    public class Data
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string ImagePath { get; set; }
        public string ImageBase64 { get; set; }

        public Data(string filename, string filepath, string imagepath, string imagebase64)
        {
            FileName = filename;
            FilePath = filepath;
            ImagePath = imagepath;
            ImageBase64 = imagebase64;
        }
    }
}
