namespace Non_visual_components_Kouvshinoff.InfoModels
{
    public class Range
    {
        /// <summary>
        /// название диапазона
        /// </summary>
        public string name;
        /// <summary>
        /// данные диапазона: (название точки, величина в точке)
        /// </summary>
        public Dictionary<string, int> data;
        /// <summary>
        /// диапазон данных
        /// </summary>
        /// <param name="name">название диапазона</param>
        /// <param name="data">данные диапазона: (название точки, величина в точке)</param>
        public Range(string name, Dictionary<string, int> data)
        {
            this.name = name;
            this.data = data;
        }
    }
}
