using System.ComponentModel;

namespace Non_visual_components_Kouvshinoff
{
    public partial class CustomComponentExcelTableWithHeader : Component
    {
        public CustomComponentExcelTableWithHeader()
        {
            InitializeComponent();
        }

        public CustomComponentExcelTableWithHeader(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }
    }
}
