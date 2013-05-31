using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WirecastWrapper;
using System.Threading;
using System.IO;
using System.Runtime.InteropServices;

namespace WirecastWrapperTest
{
    [TestClass]
    public class WirecastWrapperTesting
    {

        private const string TEST_CLIP = @"c:\Temp\screen.capture.clips\rotating.clips\clip1.xesc";
        private const string TEST_CLIP_NAME = "clip1.xesc";
        private const string BLANK_SHOT = "Blank Shot";


        private static WirecastDocument document
        {
            get
            {
                return Wirecast.Instance.DocumentByIndex(1);
            }
        }

        public static WirecastLayer layer
        {
            get
            {
                return document.DefaultLayer;
            }
        }

        [TestMethod]
        public void StartWireCast()
        {
            Wirecast.StartWirecast();

            try
            {
                var obj = Marshal.GetActiveObject("Wirecast.Application");
                Assert.AreEqual(true, true);
            }
            catch
            {
                Assert.AreEqual(true, false);
            }
        }

        [TestMethod]
        public void ActiveTransitionIndex()
        {

            document.ActiveTransitionIndex = WirecastWrapper.ActiveTransitionIndex.FirstPopup;
            Thread.Sleep(1000);
            Assert.AreEqual(WirecastWrapper.ActiveTransitionIndex.Null, document.ActiveTransitionIndex);

            document.ActiveTransitionIndex = WirecastWrapper.ActiveTransitionIndex.SecondPopup;
            Thread.Sleep(1000);
            Assert.AreEqual(WirecastWrapper.ActiveTransitionIndex.FirstPopup, document.ActiveTransitionIndex);

            document.ActiveTransitionIndex = WirecastWrapper.ActiveTransitionIndex.ThirdPopup;
            Thread.Sleep(1000);
            Assert.AreEqual(WirecastWrapper.ActiveTransitionIndex.SecondPopup, document.ActiveTransitionIndex);

            document.ActiveTransitionIndex = WirecastWrapper.ActiveTransitionIndex.FirstPopup;
            Thread.Sleep(1000);
            Assert.AreEqual(WirecastWrapper.ActiveTransitionIndex.Null, document.ActiveTransitionIndex);


        }

        [TestMethod]
        public void AudioMutedToSpeaker()
        {

            document.AudioMutedToSpeaker = true;
            Assert.AreEqual(true, document.AudioMutedToSpeaker);

            document.AudioMutedToSpeaker = false;
            Assert.AreEqual(false, document.AudioMutedToSpeaker);


        }

        [TestMethod]
        public void AutoLive()
        {

            document.AutoLive = false;
            Assert.AreEqual(false, document.AutoLive);

            document.AutoLive = true;
            Assert.AreEqual(true, document.AutoLive);
        }

        [TestMethod]
        public void IsArchivingToDisk()
        {
            Assert.AreEqual(false, document.IsArchivingToDisk);

        }

        [TestMethod]
        public void IsBroadcasting()
        {
            Assert.AreEqual(false, document.IsBroadcasting);
        }

        [TestMethod]
        public void TransitionSpeed()
        {
            Assert.AreEqual(WirecastWrapper.TransitionSpeed.normal, document.TransitionSpeed);

            document.TransitionSpeed = WirecastWrapper.TransitionSpeed.slowest;
            Assert.AreEqual(WirecastWrapper.TransitionSpeed.slowest, document.TransitionSpeed);

            document.TransitionSpeed = WirecastWrapper.TransitionSpeed.slow;
            Assert.AreEqual(WirecastWrapper.TransitionSpeed.slow, document.TransitionSpeed);

            document.TransitionSpeed = WirecastWrapper.TransitionSpeed.fast;
            Assert.AreEqual(WirecastWrapper.TransitionSpeed.fast, document.TransitionSpeed);

            document.TransitionSpeed = WirecastWrapper.TransitionSpeed.fastest;
            Assert.AreEqual(WirecastWrapper.TransitionSpeed.fastest, document.TransitionSpeed);

            document.TransitionSpeed = WirecastWrapper.TransitionSpeed.normal;
            Assert.AreEqual(WirecastWrapper.TransitionSpeed.normal, document.TransitionSpeed);
        }

        [TestMethod]
        public void ArchiveToDisk()
        {
            document.ArchiveToDiskStart();
            Thread.Sleep(5000);
            Assert.AreEqual(true, document.IsArchivingToDisk);

            document.ArchiveToDiskStop();
            Thread.Sleep(2000);
            Assert.AreEqual(false, document.IsArchivingToDisk);

        }

        [TestMethod]
        public void Broadcast()
        {
            document.BroadcastStart();
            Thread.Sleep(500);
            Assert.AreEqual(true, document.IsBroadcasting);


            document.BroadcastStop();
            Thread.Sleep(1000);
            Assert.AreEqual(false, document.IsBroadcasting);

        }

        // test failed
        [TestMethod]
        public void SaveSnapshot()
        {
            /*
            string filename = Environment.SpecialFolder.MyDocuments + "\\sampleshot";
            document.SaveSnapshot( filename );
            Thread.Sleep(2000);
            Assert.AreEqual(true,File.Exists( filename ));
            */
        }

        [TestMethod]
        public void LayerByIndex()
        {
            var l = document.LayerByOneBasedIndex(5);
            Assert.IsNotNull(l);
            Assert.AreEqual(1, l.ShotCount);
        }

        [TestMethod]
        public void LayerByName()
        {
            for (int i = 1; i <= 5; i++)
            {
                var l = document.LayerByName("Master Layer " + i.ToString());
                Assert.IsNotNull(l);
            }
        }

        [TestMethod]
        public void ShotByName()
        {
            var shot = document.ShotByName(BLANK_SHOT, WirecastWrapper.CompareMethod.ExactMatch);
            Assert.IsNotNull(shot);
            Console.WriteLine(shot.ShotID);

        }

        [TestMethod]
        public void ShotByID()
        {
            var shot_named = document.ShotByName(BLANK_SHOT, WirecastWrapper.CompareMethod.ExactMatch);
            var shot_ided = document.ShotByShotID(shot_named.ShotID);

            Assert.AreEqual(shot_named, shot_ided);
        }

        [TestMethod]
        public void AddShotWithMedia()
        {

            layer.AddShotWithMedia(TEST_CLIP);

        }

        [TestMethod]
        public void RemoveMedia()
        {

            layer.AddShotWithMedia(TEST_CLIP);

            var afterAddShotCount = layer.ShotCount;

            Assert.AreEqual(afterAddShotCount, layer.ShotCount);

            document.RemoveMedia(TEST_CLIP);

            // Wirecast does not remove the shot, just the media!
            Assert.AreEqual(afterAddShotCount, layer.ShotCount);
        }

        [TestMethod]
        public void AddRemoveShots()
        {


            var currentShots = layer.ShotCount;
            layer.AddShotWithMedia(TEST_CLIP);

            Assert.AreEqual(currentShots + 1, layer.ShotCount);

            var shot = layer.GetShotByName(TEST_CLIP_NAME, WirecastWrapper.CompareMethod.ExactMatch);
            Assert.AreEqual(TEST_CLIP_NAME, shot.Name);

            layer.RemoveShot(shot);

            Assert.AreEqual(currentShots, layer.ShotCount);

        }

        [TestMethod]
        public void ShotsList()
        {
            Assert.AreEqual(layer.Shots.Count, layer.ShotCount);

        }

        [TestMethod]
        public void ActiveShot()
        {

            layer.ActiveShot = layer.GetShotByName(BLANK_SHOT);
            Assert.AreEqual("Blank Shot", layer.ActiveShot.Name);
        }

        [TestMethod]
        public void ShotByIndexAndName()
        {
            var shot_id1 = layer.ShotIDByIndex(1);
            var shot_id2 = layer.ShotIDByName("Blank Shot");
            Assert.AreEqual(shot_id1, shot_id2);

            var shot1 = layer.GetShotByID(shot_id1);
            var shot2 = layer.GetShotByID(shot_id2);
            Assert.AreEqual(shot1, shot2);
        }

        [TestMethod]
        public void Go()
        {
            layer.Go();
        }

        [TestMethod]
        public void WirecastDocumentAndDefaultLayer()
        {
            Assert.AreEqual(layer.Document.LayerByOneBasedIndex(Wirecast.DEFAULT_LAYER_INDEX + 1), layer);
            Assert.AreEqual(layer.Document.Layers[Wirecast.DEFAULT_LAYER_INDEX], layer);
            Assert.AreEqual(layer.Document.DefaultLayer, layer);
        }

        [TestMethod]
        public void Visible()
        {
            layer.Visible = false;
            Assert.AreEqual(false, layer.Visible);

            layer.Visible = true;
            Assert.AreEqual(true, layer.Visible);
        }

        [TestMethod]
        public void Layers()
        {
            int i = 0;
            foreach (var layer in document.Layers)
            {
                i++;
                Assert.AreEqual(layer, document.LayerByOneBasedIndex(i));
            }
        }

        [TestMethod]
        public void ShotThruLayer()
        {

            Assert.AreEqual("Blank Shot", document.Layers[0].Shots[0].Name);

        }



    }
}

