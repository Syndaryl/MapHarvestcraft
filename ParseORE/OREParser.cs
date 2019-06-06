using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ParseORE
{
    public class OREParser
    {
        public StreamReader Reader { get; private set; }
        public bool EOF { get {
                return Reader.EndOfStream;
            } }
        public readonly Dictionary<string, Item> Items;

        public Dictionary<string, OreBlock> Ores { get; private set; } = new Dictionary<string, OreBlock>();

        //public string Ore { get; private set; }
        public OREParser(string path)
        {
            Reader = new StreamReader(path);
        }

        public OREParser(string path, Encoding encoding, Dictionary<string, Item> items)
        {
            Reader = new StreamReader(path, encoding);
            Items = items;
        }

        public void ReadAllOres()
        {
            while (!EOF)
            {
                NextOreBlock();
            }
        }

        public OreBlock NextOreBlock()
        {
            string name = "";
            List<Item> members = new List<Item>();

            if (Reader.EndOfStream) throw new EndOfStreamException();

            var line =  Reader.ReadLine();
            name = ExtractItem(line);

            while (!Reader.EndOfStream)
            {
                if (Reader.Peek() == '-')
                {
                    line = ExtractItem(Reader.ReadLine());
                    var metaIDX = line.LastIndexOf(':');
                    string meta="";
                    if (line.IndexOf(':') != metaIDX)
                    {
                        // two colons = hasmeta
                        meta = line.Substring(metaIDX + 1);
                        line = line.Substring(metaIDX - 1);
                    }

                    Item baseItem = new Item { IDName = line, Metadata = meta };
                    if (Items.ContainsKey(line))
                    {
                        baseItem = Items[line];
                    }
                    else
                    {
                        Items[line] = baseItem;
                    }
                    members.Add(item: baseItem);
                }
                else
                {
                    break;
                }
            }

            var oreBlock = new OreBlock(name, members.ToArray());
            Ores.Add(name, oreBlock);
            return oreBlock;
        }

        private static string ExtractItem(string line)
        {
            string name="";
            int start = -1;
            int end = -1;
            int length = -1;
            try
            {
                start = line.IndexOf('<') + 1;
                end = line.IndexOf('>');
                length = end - start;
                name = line.Substring(start, length);
            }
            catch (ArgumentOutOfRangeException e)
            {
                Console.Error.WriteLine("ArgumentOutOfRange exception thrown when processing '{0}';\nbrokets found to start at {2} and end at {3} for a length of {4}\n{1}", line, e.ToString(), start, end, length);
            }
            catch (IndexOutOfRangeException e)
            {
                Console.Error.WriteLine("IndexOutOfRange exception thrown when processing '{0}'\n{1}", line, e.ToString());
            }
            return name;
        }
    }
}
