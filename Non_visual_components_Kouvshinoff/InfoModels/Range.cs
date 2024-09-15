namespace Non_visual_components_Kouvshinoff.InfoModels
{
    public class Range
    {
        public string name;
        public Dictionary<string, int> data;
        public Range(string name, Dictionary<string, int> data)
        {
            this.name = name;
            this.data = data;
        }
    }
}
