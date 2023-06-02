using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using DocumentFormat.OpenXml;
using System;


namespace Powerpoint_Import
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\test.pptx";
            // Open the presentation as read-only.
            GetShapesSlide("C:\\1.pptx", 0);
            Console.ReadKey();   
        }
        // Get all the text in a slide.
        public static void GetShapesSlide(string presentationFile, int slideIndex)
        {
            // Open the presentation as read-only.
            using (PresentationDocument ppt = PresentationDocument.Open(presentationFile, true))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                for(var i= 0; i < part.SlideParts.Count(); i++)
                {
                    string relId = (slideIds[i] as SlideId).RelationshipId;
                    // Get the slide part from the relationship ID.
                    SlidePart slide = (SlidePart)part.GetPartById(relId);
                    if (slide != null)
                    {
                        // Get the shape tree that contains the shape to change.
                        var shapeTree = slide.Slide.CommonSlideData.ShapeTree;
                        // Get the first shape in the shape tree.

                        if (shapeTree.ChildElements.Count > 0)
                        {
                            foreach (var child in shapeTree.ChildElements)
                            {
                                GetShapeTreeChildrens(child);
                            }
                        }
                    }
                }
            }
        }

        private static void GetShapeTreeChildrens(OpenXmlElement child)
        {
            switch (child.LocalName)
            {
                case "grpSp":
                    HandleGrouping(child);
                    break;
                case "sp":
                    ProccessShape(child);
                    break;
                case "cxnSp":
                    ProccessShape(child);
                    break;
                default:
                    Console.WriteLine($"object omited: {child.LocalName}");
                    break;
            }
        }

        private static void HandleGrouping(OpenXmlElement group)
        {
            foreach (var child in group.ChildElements)
            {
                if (child.LocalName == "sp")
                {
                    ProccessShape(child);
                }
                else if (child.LocalName == "grpSp")
                {
                    HandleGrouping(child);
                }
            }
        }

        private static void ProccessShape(OpenXmlElement child)
        {
            Console.WriteLine($" Shape/Line Start ---------------------------------");
            Console.WriteLine(child.InnerText);
            Console.WriteLine(child.InnerXml);
            Console.WriteLine($" Shape/Line End ---------------------------------");
        }
    }
}
