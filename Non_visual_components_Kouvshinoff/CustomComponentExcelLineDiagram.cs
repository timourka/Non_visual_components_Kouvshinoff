using System.ComponentModel;

namespace Non_visual_components_Kouvshinoff
{
    public partial class CustomComponentExcelLineDiagram : Component
    {
        public CustomComponentExcelLineDiagram()
        {
            InitializeComponent();
        }

        public CustomComponentExcelLineDiagram(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }
    }
}
