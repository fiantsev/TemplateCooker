using System;
using System.Collections.Generic;
using System.Linq;

namespace TemplateCooking.Service.OperationExecutors
{
    public static class ListExtensions
    {
        public static List<List<T>> Transpose<T>(this List<List<T>> originalMatrix)
        {
            if (originalMatrix.Count == 0)
                return new List<List<T>>();

            var originalRowCount = originalMatrix.Count;
            var originalColumnCount = originalMatrix[0].Count;

            if (originalMatrix.Any(row => row.Count != originalColumnCount))
                throw new Exception("Количество элементов в каждой строке должно быть одинаковым");

            var resultMatrix = new List<List<T>>(originalColumnCount);

            for (var columnIndex = 0; columnIndex < originalColumnCount; ++columnIndex)
            {
                resultMatrix.Add(new List<T>(originalRowCount));
                for (var rowIndex = 0; rowIndex < originalRowCount; ++rowIndex)
                    resultMatrix[columnIndex].Add(originalMatrix[rowIndex][columnIndex]);
            }

            return resultMatrix;
        }
    }
}
