using Labo.Common.Patterns;

namespace Labo.Excel
{
    /// <summary>
    /// The excel reader writer factory.
    /// </summary>
    public sealed class ExcelReaderWriterFactory : Factory<ExcelReaderWriterType, IExcelReaderWriter>
    {
        private static readonly ExcelReaderWriterFactory s_Instance = new ExcelReaderWriterFactory();

        /// <summary>
        /// Prevents a default instance of the <see cref="ExcelReaderWriterFactory"/> class from being created.
        /// </summary>
        private ExcelReaderWriterFactory()
        {
            RegisterProvider(ExcelReaderWriterType.NPOI, () => new NPoiExcelReaderWriter(), true);
            RegisterProvider(ExcelReaderWriterType.OFFICEEXCELINTEROP, () => new OfficeInteropExcelReaderWriter());
        }

        /// <summary>
        /// Gets the instance.
        /// </summary>
        /// <value>
        /// The instance.
        /// </value>
        public static ExcelReaderWriterFactory Instance
        {
            get { return s_Instance; }
        }
    }
}
