namespace ParseORE
{
    using System;
    using System.Collections;
    using System.Collections.Generic;

    public class OreBlock : IEnumerable<Item>, ICollection<Item>, IDictionary<string, Item>
    {
        public string Name {get; private set;}
        public Dictionary<string, Item> Members { get; private set; } = new Dictionary<string, Item>();

        public OreBlock(string name, Item[] members)
        {
            Name = name;
            foreach (var item in members)
            {
                if ( !Members.ContainsKey(item.IDName) )
                    Members.Add(item.IDName, item);
            }
        }

        #region Interfaces

        public int Count
        {
            get
            {
                return ((ICollection<Item>)Members).Count;
            }
        }

        public bool IsReadOnly
        {
            get
            {
                return ((ICollection<Item>)Members).IsReadOnly;
            }
        }

        public ICollection<string> Keys => ((IDictionary<string, Item>)Members).Keys;

        public ICollection<Item> Values => ((IDictionary<string, Item>)Members).Values;

        public Item this[string key] { get => ((IDictionary<string, Item>)Members)[key]; set => ((IDictionary<string, Item>)Members)[key] = value; }

        public IEnumerator<Item> GetEnumerator()
        {
            return ((IEnumerable<Item>)Members).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<Item>)Members).GetEnumerator();
        }

        public void Add(Item item)
        {
            ((ICollection<Item>)Members).Add(item);
        }

        public void Clear()
        {
            ((ICollection<Item>)Members).Clear();
        }

        public bool Contains(Item item)
        {
            return ((ICollection<Item>)Members).Contains(item);
        }

        public void CopyTo(Item[] array, int arrayIndex)
        {
            ((ICollection<Item>)Members).CopyTo(array, arrayIndex);
        }

        public bool Remove(Item item)
        {
            return ((ICollection<Item>)Members).Remove(item);
        }

        public bool ContainsKey(string key)
        {
            return ((IDictionary<string, Item>)Members).ContainsKey(key);
        }

        public void Add(string key, Item value)
        {
            ((IDictionary<string, Item>)Members).Add(key, value);
        }

        public bool Remove(string key)
        {
            return ((IDictionary<string, Item>)Members).Remove(key);
        }

        public bool TryGetValue(string key, out Item value)
        {
            return ((IDictionary<string, Item>)Members).TryGetValue(key, out value);
        }

        public void Add(KeyValuePair<string, Item> item)
        {
            ((IDictionary<string, Item>)Members).Add(item);
        }

        public bool Contains(KeyValuePair<string, Item> item)
        {
            return ((IDictionary<string, Item>)Members).Contains(item);
        }

        public void CopyTo(KeyValuePair<string, Item>[] array, int arrayIndex)
        {
            ((IDictionary<string, Item>)Members).CopyTo(array, arrayIndex);
        }

        public bool Remove(KeyValuePair<string, Item> item)
        {
            return ((IDictionary<string, Item>)Members).Remove(item);
        }

        IEnumerator<KeyValuePair<string, Item>> IEnumerable<KeyValuePair<string, Item>>.GetEnumerator()
        {
            return ((IDictionary<string, Item>)Members).GetEnumerator();
        }
        #endregion Interfaces
    }

    public class Item : IIDName, IComparable, IEquatable<string>
    {
        public Item()
        {
        }

        public Item(string iDName, string metadata="0")
        {
            IDName = iDName ?? throw new ArgumentNullException(nameof(iDName));
            Metadata = metadata ?? throw new ArgumentNullException(nameof(metadata));
        }

        public string IDName { get; set; }
        public string Metadata { get; set; }

        public int CompareTo(object obj) => IDName.CompareTo(obj);

        public bool Equals(string other) => IDName.Equals(other);
    }

    internal interface IIDName
    {
        string IDName { get; set; }
    }

    public enum RecipeType
    {
        Shapeless, Shaped, Furnace
    }

    public class ItemStack : IComparable, IEquatable<string>
    {
        public Item Item { get; set; }
        public byte Count { get; set; }

        public int CompareTo(object obj)
        {
            return Item.CompareTo(obj);
        }

        public bool Equals(string other)
        {
            return Item.Equals(other);
        }
    }
}