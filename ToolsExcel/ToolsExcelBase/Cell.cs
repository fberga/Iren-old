
namespace Iren.ToolsExcel.Base
{
    public class Cell
    {
        private WidthClass _w;
        private HeightClass _h;

        public Cell()
        {
            _w = new WidthClass();
            _h = new HeightClass();
        }

        public WidthClass Width { get { return _w; } }
        public HeightClass Height { get { return _h; } }

        public class WidthClass
        {
            public double empty = 1,
            dato = 8.8,
            entita = 3,
            informazione = 28,
            unitaMisura = 6,
            parametro = 8.8,
            riepilogo = 9;            
        }

        public class HeightClass
        {
            public double normal = 15,
            empty = 5;
        }
    }
}
