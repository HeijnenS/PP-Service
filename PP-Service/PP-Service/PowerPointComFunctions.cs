using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PP_Service
{
    class PowerPointComFunctions
    {
            public string PowerpointPath { get; set; }

            PowerPoint.Application oPowerPoint = null;
            PowerPoint.Presentations oPres = null;
            PowerPoint.Presentation oPre = null;
            PowerPoint.Slides oSlides = null;
            PowerPoint.Slide oSlide = null;
            PowerPoint.Shapes oShapes = null;
            PowerPoint.Shape oShape = null;
            PowerPoint.TextFrame oTxtFrame = null;
            PowerPoint.TextRange oTxtRange = null;

            public PowerPointComFunctions()
            {
            }

            ~PowerPointComFunctions()
            {
                if (oTxtRange != null)
                {
                    Marshal.FinalReleaseComObject(oTxtRange);
                    oTxtRange = null;
                }
                if (oTxtFrame != null)
                {
                    Marshal.FinalReleaseComObject(oTxtFrame);
                    oTxtFrame = null;
                }
                if (oShape != null)
                {
                    Marshal.FinalReleaseComObject(oShape);
                    oShape = null;
                }
                if (oShapes != null)
                {
                    Marshal.FinalReleaseComObject(oShapes);
                    oShapes = null;
                }
                if (oSlide != null)
                {
                    Marshal.FinalReleaseComObject(oSlide);
                    oSlide = null;
                }
                if (oSlides != null)
                {
                    Marshal.FinalReleaseComObject(oSlides);
                    oSlides = null;
                }
                if (oPre != null)
                {
                    Marshal.FinalReleaseComObject(oPre);
                    oPre = null;
                }
                if (oPres != null)
                {
                    Marshal.FinalReleaseComObject(oPres);
                    oPres = null;
                }
                if (oPowerPoint != null)
                {
                    Marshal.FinalReleaseComObject(oPowerPoint);
                    oPowerPoint = null;
                }
            }


            public string PresentationElapsedTime()
            {
                if (hasActivePresentation())
                {
                    try
                    {
                        // in seconds  
                        float totalinsec = oPowerPoint.ActivePresentation.SlideShowWindow.View.PresentationElapsedTime;
                        int hours = Convert.ToInt16(totalinsec / 3600);
                        int minutes = Convert.ToInt16((totalinsec - (hours * 3600)) / 60);
                        int seconds = Convert.ToInt16(totalinsec - (hours * 3600) - (minutes * 60));
                        return "presentationelapsedtime-" + Convert.ToString(hours) + "h" + Convert.ToString(minutes) + "m" + Convert.ToString(seconds) + "s";
                    }
                    catch (Exception ex)
                    {
                        return "error-" + ex.Message.ToString();
                    }
                }
                else
                {
                    return "presentationelapsedtime-please select a presentation";
                }
            }

            public string SlideElapsedTime()
            {
                if (hasActivePresentation())
                {
                    try
                    {
                        // in seconds
                        float totalinsec = oPowerPoint.ActivePresentation.SlideShowWindow.View.SlideElapsedTime;
                        int hours = Convert.ToInt16(totalinsec / 3600);
                        int minutes = Convert.ToInt16((totalinsec-(hours*3600))/60);
                        int seconds = Convert.ToInt16(totalinsec - (hours * 3600) - (minutes * 60));
                        return "SlideElapsedTime-" + Convert.ToString(hours) + "h" + Convert.ToString(minutes) + "m" + Convert.ToString(seconds) + "s";
                    }
                    catch (Exception ex)
                    {
                        return "error-" + ex.Message.ToString();
                    }
                }
                else
                {
                    return "SlideElapsedTime-please select a presentation";
                }
            }

            private void test ()
            {
                oPowerPoint.ActivePresentation.SlideMaster.TimeLine.MainSequence.ToString();
            }

            public string ActiveSlide()
            {
                if (hasActivePresentation())
                {
                    try
                    {
                        return "currentslide-" + oPowerPoint.ActivePresentation.SlideShowWindow.View.Slide.SlideIndex.ToString() + "," + oPowerPoint.ActivePresentation.SlideShowWindow.View.Slide.Name;
                    }
                    catch (Exception ex)
                    {
                        return "error-" + ex.Message.ToString();
                    }
                }
                else
                {
                    return "currentslide-please select a presentation";
                }
            }

            public string SlideName(string index)
            {
                if (hasActivePresentation())
                {
                    try
                    {
                        int intIndex = Convert.ToInt16(index);
                        return "slidename-" + index + "," + oPowerPoint.ActivePresentation.Slides._Index(intIndex).Name;
                        //return "slidename-" + index.toString() + "," + oPowerPoint.ActivePresentation.Slides(index).Name;
                    }
                    catch (Exception ex)
                    {
                        return "error-" + ex.Message.ToString();
                    }
                }
                else
                {
                    return "slidename-please select a presentation";
                }
            }


            public string SlideCount()
            {
                if (hasActivePresentation())
                {
                    try
                    {
                        return "slides-" + oPowerPoint.ActivePresentation.Slides.Count.ToString();
                    }
                    catch (Exception ex)
                    {
                        return "error-" + ex.Message.ToString();
                    }
                }
                else
                {
                    return "slides-please select a presentation";
                }
            }



            public string HasActivePresentation()
            {
                try
                {
                    string tmpName = oPowerPoint.ActivePresentation.Name;
                    return "activepresentation-" + tmpName;
                }
                catch
                {
                    return "activepresentation-none";
                }
            }


            public string GoToSlide(string slidenumber)
            {
                if (hasActivePresentation())
                {
                    try
                    {
                        int tmpSlide = Convert.ToInt16(slidenumber);
                        oPowerPoint.ActivePresentation.SlideShowWindow.View.GotoSlide(tmpSlide);
                        return "gotoslide-" + slidenumber.ToString();
                    }
                    catch (Exception ex)
                    {
                        return "error-" + ex.Message.ToString();
                    }
                }
                else
                {
                    return "gotoslide-please select a presentation";
                }
            }

            private bool hasActivePresentation()
            {
                try
                {
                    string tmpName = oPowerPoint.ActivePresentation.Name;
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            public string FirstSlide()
            {
                if (hasActivePresentation())
                {
                    try
                    {
                        oPowerPoint.ActivePresentation.SlideShowWindow.View.GotoSlide(1);
                        return "firstslide";
                    }
                    catch (Exception ex)
                    {
                        return "error-" + ex.Message.ToString();
                    }
                }
                else
                {
                    return "gotoslide-please select a presentation";
                }
            }

            public string LastSlide()
            {
                if (hasActivePresentation())
                {
                    try
                    {
                        oPowerPoint.ActivePresentation.SlideShowWindow.View.GotoSlide(oPowerPoint.ActivePresentation.Slides.Count);
                        return "lastslide";
                    }
                    catch (Exception ex)
                    {
                        return "error-" + ex.Message.ToString();
                    }
                }
                else
                {
                    return "LastSlide-please select a presentation";
                }
            }

            public string NextSlide()
            {
                if (hasActivePresentation())
                {
                    try
                    {
                        oPowerPoint.ActivePresentation.SlideShowWindow.View.Next();
                        return "Next";
                    }
                    catch (Exception ex)
                    {
                        return "error-" + ex.Message.ToString();
                    }
                }
                else
                {
                    return "NextSlide-please select a presentation";
                }
            }

            public string PreviousSlide()
            {
                if (hasActivePresentation())
                {
                    try
                    {
                        oPowerPoint.ActivePresentation.SlideShowWindow.View.Previous();
                        return "previous";
                    }
                    catch (Exception ex)
                    {
                        return "error-" + ex.Message.ToString();
                    }
                }
                else
                {
                    return "PreviousSlide-please select a presentation";
                }
            }

            public string ClosePowerPoint()
            {
                if (hasActivePresentation())
                {
                    try
                    {
                        oPowerPoint.ActivePresentation.Close();
                        return "close";
                    }
                    catch (Exception ex)
                    {
                        return "error-" + ex.Message.ToString();
                    }
                }
                else
                {
                    return "ClosePowerPoint-please select a presentation";
                }
            }

            public string CloseApplication()
            {
                try
                {
                    oPowerPoint.Quit();
                    return "Application is closed";
                }
                catch (Exception ex)
                {
                    return "error-" + ex.Message.ToString();
                }                
            }

            public string OpenPowerPoint(string FileName)
            {
                try
                {
                    try
                    {
                        string tmpName = oPowerPoint.ActivePresentation.Name;
                        oPowerPoint.ActivePresentation.Close();
                    }
                    catch
                    {
                        //do nothing there's no open presentation
                    }

                    // Create an instance of Microsoft PowerPoint and make it invisible. 
                    oPowerPoint = new PowerPoint.Application();
                    // By default PowerPoint is invisible, till you make it visible. oPowerPoint.Visible = Office.MsoTriState.msoFalse; 
                    // Create a new Presentation. 
                    if(File.Exists(PowerpointPath+ "\\" + FileName))
                    {
                        oPowerPoint.Presentations.Open(PowerpointPath+ "\\" + FileName);
                        oPowerPoint.ActivePresentation.SlideShowSettings.Run();
                        return "open-" + PowerpointPath + "\\" + FileName;
                    }
                    else
                    {
                        return "open-file [" + PowerpointPath+ "\\" + FileName + "] does not exist, check ?powerpoints";
                    }
                }
                catch (Exception ex)
                {
                    return "error-"+ex.Message.ToString();
                }
            }
    }
}
