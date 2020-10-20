using TemplateCooker.Domain.Markers;

namespace TemplateCooker.Domain.LayoutShifts
{
    public abstract class LayoutShift
    {
        /// <summary>
        /// опорная позиция - позиция с которой необходимо выполнять смещение
        /// </summary>
        public MarkerPosition OriginPosition { get; set; }

        /// <summary>
        /// количественное значение смещения
        /// </summary>
        public int Amount { get; set; }
    }

    /// <summary>
    /// смещение строк целиком
    /// </summary>
    public class RowLayoutShift : LayoutShift
    {

    }
}