/*
 Wirecast 4.3 .NET Wrapper v 0.1 by Kristijan Burnik <kristijanburnik@gmail.com>
*/

using System;
using System.Collections.Generic;
using System.Text;

namespace WirecastWrapper
{
    public class WirecastShot
    {
        public int ShotID { get; set; }
        public string Name { get; set; }

        private static Dictionary<int, WirecastShot> shots = new Dictionary<int, WirecastShot>();

        public static WirecastShot CreateFromObject(int shotID, object shot)
        {
            string name = (string)Late.Get(shot, "Name");
            if (!shots.ContainsKey(shotID))
            {
                shots[shotID] = new WirecastShot() { ShotID = shotID, Name = name };
            }

            return shots[shotID];

        }
    }

}
