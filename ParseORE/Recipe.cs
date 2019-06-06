namespace ParseORE
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    public class BenchFoodRecipe : IRecipe<OreBlock>, IIDName, IDisposable, IComparable, IEquatable<string>
    {

        public string IDName { get; set; } = "";
        public RecipeType Type { get; set; } = RecipeType.Shapeless;
        public List<OreBlock> Inputs { get; set; } = new List<OreBlock>(9);
        public ItemStack Output { get; set; } = new ItemStack();
        public double FoodRestore { get; set; } = 0;
        public double Saturation { get {
                return Inputs.Count * FoodRestore==0? 0: 1;
            }
        }

        public static BenchFoodRecipe BenchFoodRecipeFactory(List<BenchFoodRecipe> recipes, List<string> row, Dictionary<string, OreBlock> oreDictionary, Dictionary<string, Item> items)
        {
            //      1             2            3             4       5     6             7            8          9           10          11          12          13          14
            //ID Name	Recipe Type	Food Restore	Saturation	Amount 	Tool	Ingredient	Ingredient2	Ingredient3	Ingredient4	Ingredient5	Ingredient6	Ingredient7	Ingredient8
            var name = row[1];
            int idx = 0;
            name = string.Format("{0} {1}", row[1], ++idx);
            while (recipes.Find(r => r.IDName.Equals(name)) != null)
            {
                name = string.Format("{0} {1}", row[1], ++idx);
            }
            double outCount = 1;
            if (!string.IsNullOrEmpty(row[5]))
            {
                outCount = double.Parse(row[5]);
            }
            Item baseItem = new Item(row[1]);
            if (items.ContainsKey(row[1]))
            {
                baseItem = items[row[1]];
            } else
            {
                items[row[1]] = baseItem;
            }
            var outputItemStack = new ItemStack { Item = baseItem, Count = (byte)Math.Round(outCount) };

            List<OreBlock> ins = new List<OreBlock>(9);
            for (int i = 6; i < row.Count; i++)
            {
                if (!string.IsNullOrEmpty(row[i]))
                {
                    if (oreDictionary.ContainsKey(row[i]))
                    {
                        ins.Add(oreDictionary[row[i]]);
                    }
                }
            }

            bool success = double.TryParse(row[3], out double restoration);
            if (!success) restoration = 0;
            var recipe = new BenchFoodRecipe
            {
                IDName = name,
                Type = RecipeType.Shapeless,
                FoodRestore = restoration,
                Output = outputItemStack,
                Inputs = ins
            };
            return recipe;
        }

        #region Constructors
        public BenchFoodRecipe()
        {
        }

        public BenchFoodRecipe(string iDName, RecipeType type, double foodRestore, List<OreBlock> inputs, ItemStack output)
        {
            IDName = iDName ?? throw new ArgumentNullException(nameof(iDName));
            Type = type;
            Inputs = inputs ?? throw new ArgumentNullException(nameof(inputs));
            Output = output ?? throw new ArgumentNullException(nameof(output));
            FoodRestore = foodRestore;
        }

        #endregion Constructors

        #region Interface Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Inputs.RemoveAll(x => true);
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~BenchFoodRecipe() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }

        public int CompareTo(object obj)
        {
            return ((IComparable)Output).CompareTo(obj);
        }

        public bool Equals(string other)
        {
            return ((IEquatable<string>)Output).Equals(other);
        }
        #endregion

    }
    public class BenchRecipe : IRecipe<ItemStack>, IIDName, IDisposable, IComparable, IEquatable<string>
    {
        public string IDName { get; set; } = "";
        public RecipeType Type { get; set; } = RecipeType.Shapeless;
        public List<ItemStack> Inputs { get; set; } = new List<ItemStack>(9);
        public ItemStack Output { get; set; } = new ItemStack();

        #region Interface Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Inputs.RemoveAll(x => true);
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~BenchFoodRecipe() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }

        public int CompareTo(object obj)
        {
            return IDName.CompareTo(obj);
        }

        public bool Equals(string other)
        {
            return IDName.Equals(other);
        }
        #endregion

    }

    interface IRecipe<T>
    {
        RecipeType Type { get; set; }
        List<T> Inputs { get; set; }
        ItemStack Output { get; set; }
    }
}
