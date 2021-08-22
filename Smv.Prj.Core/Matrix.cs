using System;
using System.Collections.Generic;
using System.Text;



namespace Smv.Collections.Matrix
{
  /*
  public static class Program
  {
    private static void Main()
    {
      var ints = new[,] { { 1, 2, 3 }, { 4, 5, 6 } };
      var matrix = new Matrix<int>(ints);
      matrix.AddColumn(1);
      matrix.AddRow(1);
      Console.WriteLine(matrix.ToString());
      matrix.RemoveColumn(1);
      matrix.RemoveRow(1);
      Console.WriteLine(matrix.ToString());
      Console.ReadKey();
    }
  }
 */
  public static class ArrayUtl
  {
    public static Array ResizeArray(Array arr, int[] newSizes)
    {
      if (newSizes.Length != arr.Rank)
        throw new ArgumentException(@"arr must have the same number of dimensions as there are elements in newSizes", nameof(newSizes));

      var temp = Array.CreateInstance(arr.GetType().GetElementType(), newSizes);
      int length = arr.Length <= temp.Length ? arr.Length : temp.Length;
      Array.ConstrainedCopy(arr, 0, temp, 0, length);
      return temp;
    }
  }


  public class Matrix<T> where T : new()
  {
    private readonly List<List<T>> matrix;

    ///*<summary>
    ///*Cоздание*матрицы.
    ///*</summary>
    ///*<param*name="rowsCount">Количество*строк.</param>
    ///*<param*name="columnCount">Количество*столбцов.</param>
    public Matrix(int rowsCount = 2, int columnCount = 2)
    {
      ColumnCount = columnCount;
      RowsCount = rowsCount;
      matrix = new List<List<T>>(rowsCount);
      for (int i = 0; i < rowsCount; i++)
      {
        var list = new List<T>(columnCount);
        for (int j = 0; j < columnCount; j++)
        {
          list.Add(default(T));
        }
        matrix.Add(list);
      }
    }

    ///*<summary>
    ///*Cоздание*матрицы.
    ///*</summary>
    ///*<param*name="data">Исходный*двумерный*массив.</param>
    public Matrix(T[,] data)
    {
      RowsCount = data.GetLength(0);
      ColumnCount = data.GetLength(1);
      matrix = new List<List<T>>(RowsCount);
      for (int i = 0; i < RowsCount; i++)
      {
        var list = new List<T>(ColumnCount);
        for (int j = 0; j < ColumnCount; j++)
        {
          list.Add(data[i, j]);
        }
        matrix.Add(list);
      }
    }

    ///*<summary>
    ///*Элемент*матрицы.
    ///*</summary>
    ///*<param*name="i">Индекс*строки.</param>
    ///*<param*name="j">Индекс*столбца.</param>
    ///*<returns></returns>
    public T this[int i, int j]
    {
      get { return matrix[i][j]; }
      set { matrix[i][j] = value; }
    }

    ///*<summary>
    ///*Количество*строк.
    ///*</summary>
    public int RowsCount { get; private set; }

    ///*<summary>
    ///*Количество*столбцов.
    ///*</summary>
    public int ColumnCount { get; private set; }

    ///*<summary>
    ///*Добавить*строку.
    ///*</summary>
    ///*<param*name="index">Индекс*вставки*строки.</param>
    public void AddRow(int index)
    {
      RowsCount++;
      var list = new List<T>(ColumnCount);
      for (int j = 0; j < ColumnCount; j++)
      {
        list.Add(default(T));
      }
      matrix.Insert(index, list);
    }

    ///*<summary>
    ///*Добавить*столбец.
    ///*</summary>
    ///*<param*name="index">Индекс*вставки*столбца.</param>
    public void AddColumn(int index)
    {
      ColumnCount++;
      foreach (var list in matrix)
      {
        list.Insert(index, default(T));
      }
    }

    ///*<summary>
    ///*Удалить*строку.
    ///*</summary>
    ///*<param*name="index">Индекс*вставки*строки.</param>
    public void RemoveRow(int index)
    {
      RowsCount--;
      matrix.RemoveAt(index);
    }

    ///*<summary>
    /// Удалить столбец.
    ///*</summary>
    ///*<param*name="index">Индекс*вставки*столбца.</param>
    public void RemoveColumn(int index)
    {
      ColumnCount--;
      foreach (var list in matrix)
      {
        list.RemoveAt(index);
      }
    }

    public override string ToString()
    {
      var stringBuilder = new StringBuilder();
      for (int i = 0; i < RowsCount; i++)
      {
        for (int j = 0; j < ColumnCount; j++)
        {
          stringBuilder.Append(matrix[i][j]);
          stringBuilder.Append("*");
        }
        stringBuilder.AppendLine();
      }
      return stringBuilder.ToString();
    }
  }
}

