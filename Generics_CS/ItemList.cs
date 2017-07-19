using System.Collections.Generic;
/// <summary>
/// Class that inherits from a simple list. Demonstrates how to implement a constructor that takes a list.
/// </summary>
/// <remarks></remarks>
public class ItemList : List<string>
{
    /// <summary>
    /// Default constructor
    /// </summary>
    /// <remarks></remarks>
    public ItemList()
    {

    }
    /// <summary>
    /// We want a constructor that takes a list of strings.
    /// </summary>
    /// <param name="items"></param>
    /// <remarks></remarks>
    public ItemList(List<string> items)
    {
        foreach (string Item in items)
        {
            this.Add(Item);
        }
    }

}