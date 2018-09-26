using DocumentFormat.OpenXml;

namespace ExcelTransformation.TableClasses
{
    public static class Extensions
    {
        public static OpenXmlElement GetElementSafe(this OpenXmlElementList collection, int index)
        {
            if (index >= collection.Count)
                return null;

            return collection[index];
        }

        public static T GetElementSafe<T>(this OpenXmlElementList collection, int index)
            where T : OpenXmlElement
        {
            return (T)GetElementSafe(collection, index);
        }
    }
}
