/*
 Wirecast 4.3 .NET Wrapper v 0.1 by Kristijan Burnik <kristijanburnik@gmail.com>
*/

using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace WirecastWrapper
{
    public class WirecastLayer
    {

        private object oDocumentLayerInterface;

        private WirecastDocument _wirecastDocument;

        public WirecastDocument Document { get { return _wirecastDocument; } }

        private object Invoke(string command)
        {
            return Late.Invoke(oDocumentLayerInterface, command);
        }

        private object Invoke(string command, object parameter)
        {
            return Late.Invoke(oDocumentLayerInterface, command, parameter);
        }

        private object Invoke(string command, object parameter, object parameter2)
        {
            return Late.Invoke(oDocumentLayerInterface, command, new object[] { parameter, parameter2 });
        }

        private object Get(string property)
        {
            return Late.Get(oDocumentLayerInterface, property);
        }

        private void Set(string property, object value)
        {
            Late.Set(oDocumentLayerInterface, property, value);
        }

        ///////////////////////////////////////////////////////////////
        private static Dictionary<int, WirecastLayer> oWirecastLayers
            = new Dictionary<int, WirecastLayer>();

        public static WirecastLayer CreateFromObject(int index, object p,  WirecastDocument wirecastDocument)
        {
            if (!oWirecastLayers.ContainsKey(index))
            {
                oWirecastLayers[index] = new WirecastLayer(p,wirecastDocument);
            }
            return oWirecastLayers[index];
        }

        private static Dictionary<string, WirecastLayer> oWirecastLayersByName
            = new Dictionary<string, WirecastLayer>();

        public static WirecastLayer CreateFromObject(string name, object p, WirecastDocument wirecastDocument)
        {
            if (!oWirecastLayersByName.ContainsKey(name))
            {
                oWirecastLayersByName[name] = new WirecastLayer(p,wirecastDocument);
            }
            return oWirecastLayersByName[name];
        }

        ///////////////////////////////////////////////////////////////
        public WirecastLayer(object oDocumentLayerInterface, WirecastDocument wirecastDocument) {
            this.oDocumentLayerInterface = oDocumentLayerInterface;
            this._wirecastDocument = wirecastDocument;
        }



        public int ShotCount
        {
            get
            {
                return (int)Invoke("ShotCount");
            }
        }

        public int ShotIDByIndex(int index)
        {
            return (int)Invoke("ShotIDByIndex", index);
        }

        public int ShotIDByName(string name, CompareMethod compareMethod = CompareMethod.ExactMatch)
        {
            return (int)Invoke("ShotIDByName", name, (int)compareMethod);
        }

        public WirecastShot GetShotByOneBasedIndex(int index)
        {
            int shot_id = ShotIDByIndex(index);
            return _wirecastDocument.ShotByShotID(shot_id);
        }

        public WirecastShot GetShotByName(string name, CompareMethod compareMethod = CompareMethod.ExactMatch)
        {
            int shot_id = ShotIDByName(name, compareMethod);
            return _wirecastDocument.ShotByShotID(shot_id);
        }

        public WirecastShot GetShotByID(int ShotID)
        {
            return _wirecastDocument.ShotByShotID(ShotID);
        }

        public WirecastShot AddShotWithMedia(string filepath)
        {
            int shot_id = (int)Invoke("AddShotWithMedia", filepath);
            return _wirecastDocument.ShotByShotID(shot_id);
        }

        public void RemoveShotByID(int id)
        {
            Invoke("RemoveShotByID", id);
        }

        public void RemoveShot(WirecastShot shot)
        {
            RemoveShotByID(shot.ShotID);
        }

        public void Go()
        {
            Invoke("Go");
        }



        public bool Visible
        {
            get
            {
                return ((int)Get("Visible")) != 0;
            }
            set
            {
                Set("Visible", (int)((value) ? 1 : 0));
            }
        }

        public WirecastShot ActiveShot
        {
            get
            {
                int shot_id = (int)Get("ActiveShotID");
                if (shot_id == 0) {
                    return null;
                } else {
                    return _wirecastDocument.ShotByShotID(shot_id);
                }
            }

            set {
                
                var shot_id = (value != null) ? value.ShotID : 0;
                Set("ActiveShotID", shot_id);
            }
        }

        private BindingList<WirecastShot> _shots = new BindingList<WirecastShot>();

        public BindingList<WirecastShot> Shots { get {
            var shotCount = ShotCount;
            for (int i = 1; i <= shotCount; i++)
            {
                var shot = GetShotByOneBasedIndex(i);
                if (_shots.Count <= i)
                {
                    _shots.Add( shot );
                }
                else
                {
                    _shots[i - 1] = shot;
                }
            }
            for (int i = shotCount; i < _shots.Count; i++)
            {
                _shots.RemoveAt(i);
            }

            return _shots;
        } }


    }

}
