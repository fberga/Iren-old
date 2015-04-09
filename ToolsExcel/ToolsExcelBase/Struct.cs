﻿
namespace Iren.ToolsExcel.Base
{
    public class Struct
    {
        #region Strutture

        public struct Cella
        {
            public struct Width
            {
                public double empty,
                    dato,
                    entita,
                    informazione,
                    unitaMisura,
                    parametro,
                    jolly1,
                    riepilogo;
            }
            public struct Height
            {
                public double normal,
                    empty;
            }

            public Width width;
            public Height height;
        }

        #endregion

        #region Variabili

        public static string tipoVisualizzazione = "O";
        public static int intervalloGiorni = 0;
        public static bool visualizzaRiepilogo = true;
        public static Cella cell;

        public int colBlock = 5,
            rigaBlock = 6,
            rigaGoto = 3,
            colRecap = 165,
            rowRecap = 2;
        public bool visData0H24 = false,
            visParametro = false;

        #endregion
    }
}
