/*
 Wirecast 4.3 .NET Wrapper v 0.1 by Kristijan Burnik <kristijanburnik@gmail.com>
*/

using System;
using System.Collections.Generic;
using System.Text;

namespace WirecastWrapper
{
    public class WirecastDocument
    {

        private object oDocumentInterface;

        private List<WirecastLayer> _layers = new List<WirecastLayer>();

        public List<WirecastLayer> Layers { get { return _layers; } }

        public WirecastDocument(object oDocumentInterface)
        {
            this.oDocumentInterface = oDocumentInterface;
            for (int i = 1; i <= Wirecast.TOTAL_LAYER_COUNT; i++) {
                _layers.Add( LayerByOneBasedIndex(i) );
            }
        }

        private object Invoke(string command)
        {
            return Late.Invoke(oDocumentInterface, command);
        }

        private object Invoke(string command, object parameter)
        {
            return Late.Invoke(oDocumentInterface, command, parameter);
        }

        private object Invoke(string command, object parameter, object parameter2)
        {
            return Late.Invoke(oDocumentInterface, command, new object[] {parameter, parameter2});
        }

        private object Get(string property)
        {
            return Late.Get(oDocumentInterface, property);
        }

        private void Set(string property, object value)
        {
            Late.Set(oDocumentInterface, property, value);
        }
        ///////////////////////////////////////////////////////////////
        private static Dictionary<int, WirecastDocument> oWirecastDocuments
            = new Dictionary<int, WirecastDocument>();

        public static WirecastDocument CreateFromObject(int index, object p)
        {
            if (!oWirecastDocuments.ContainsKey(index))
            {
                oWirecastDocuments[index] = new WirecastDocument(p);
            }
            return oWirecastDocuments[index];
        }

        ///////////////////////////////////////////////////////////////

        public void BroadcastStart()
        {
            this.Invoke("Broadcast", "start");
        }

        public void BroadcastStop()
        {
            this.Invoke("Broadcast", "stop");
        }

        public void ArchiveToDiskStart()
        {
            this.Invoke("ArchiveToDisk", "start");
        }

        public void ArchiveToDiskStop()
        {
            this.Invoke("ArchiveToDisk", "stop");
        }

        
        /// <summary>
        /// Return layer by 1-based Index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public WirecastLayer LayerByOneBasedIndex(int index)
        {
            return WirecastLayer.CreateFromObject(index,Invoke("LayerByIndex",index),this);
        }
       
        public WirecastLayer LayerByName(string Name)
        {
            return WirecastLayer.CreateFromObject(Name, Invoke("LayerByName", Name),this);
        }

        
        public WirecastShot ShotByName(string name, CompareMethod compare_method)
        {
            int shot_id = (int) this.Invoke("ShotIDByName", name,(int) compare_method);
            return ShotByShotID( shot_id );
        }

        // todo: implement
        public void SaveSnapshot(string path)
        {
            Invoke("SaveSnapshot", path);
        }


        public WirecastShot ShotByShotID(int ShotID)
        {
            object shot = this.Invoke("ShotByShotID", ShotID);
            return WirecastShot.CreateFromObject(ShotID, shot);
        }

        public void RemoveMedia(string path)
        {
            this.Invoke("RemoveMedia", path);
        }


        public WirecastLayer DefaultLayer { get {
            return Layers[ Wirecast.DEFAULT_LAYER_INDEX ];
        } }

        public bool IsBroadcasting
        {
            get
            {
                return (int)this.Invoke("IsBroadcasting") != 0;
            }
        }

        public bool IsArchivingToDisk
        {
            get
            {
                return (int)this.Invoke("IsArchivingToDisk") != 0;
            }
        }

        public bool AutoLive
        {
            get
            {
                return ((int)Get("AutoLive")) != 0;
            }
            set
            {
                int val = (value) ? 1 : 0;
                Set("AutoLive", val);
            }
        }


        public ActiveTransitionIndex ActiveTransitionIndex
        {
            get
            {
                return (ActiveTransitionIndex)
                            (int) Get("ActiveTransitionIndex");
                        
            }
            set
            {
                Set("ActiveTransitionIndex", (int) value);
            }
        }


        public bool AudioMutedToSpeaker
        {
            get
            {
                return ((int)Get("AudioMutedToSpeaker")) != 0;
            }
            set
            {
                int val = (value) ? 1 : 0;
                Set("AudioMutedToSpeaker", val);
            }
        }

        public TransitionSpeed TransitionSpeed
        {
            get
            {
                return (TransitionSpeed)
                        Enum.Parse(
                            typeof(TransitionSpeed),
                            (string)Get("TransitionSpeed")
                        );
            }
            set
            {
                Set("TransitionSpeed", value.ToString());
            }
        }



    }

}
