using System;
using System.Collections.Generic;
using System.Data;

using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;

using System.Text;
using GemBox.Document;
using GemBox.Document.Drawing;
using System.IO;
using GemBox.Document.Tables;
using System.Drawing;

namespace WebApplication2
{
    public class CVDocument
    {
        public static DocumentModel LoadDocumentWithHeader(object cv, string templatePath, string docuform)
        {
            Stream picture = new MemoryStream();
            if (cv != null)
            {
                //picture = new MemoryStream(cv);
            }
            //Image image = Image.FromStream(new MemoryStream(cv.Consultant.Picture));
            //double height = image.Height;
            //int width = image.Width;

            //if (width > 100)
            //{
            //    double newhight = (double)width / 100;
            //    height = height / newhight;
            //    width = 100;
            //}
            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("");

            DocumentModel document = DocumentModel.Load(templatePath);
            document.DefaultCharacterFormat.FontName = "Lacuna Regular";
            Table[] tables = document.GetChildElements(true, ElementType.Table).Cast<Table>().ToArray();

            Table consultantTable = tables[0];
            //Insert Image
            //if (cv != null)
            //{

            //    Picture profilePicture = new Picture(document, picture, PictureFormat.Png, width, height);
            //    if (docuform == "doc")
            //    {
            //        profilePicture.Layout = new FloatingLayout(new HorizontalPosition(7.5, LengthUnit.Inch, HorizontalPositionAnchor.Page),
            //        new VerticalPosition(0, LengthUnit.Inch, VerticalPositionAnchor.Page),
            //        profilePicture.Layout.Size);
            //        consultantTable.Rows[0].Cells[1].Blocks.Add(new Paragraph(document, profilePicture));
            //    }
            //    else
            //    {
            //        consultantTable.Rows[0].Cells[1].Blocks.Add(new Paragraph(document, profilePicture)
            //        {
            //            ParagraphFormat = new ParagraphFormat()
            //            {
            //                Alignment = HorizontalAlignment.Right
            //            }
            //        });
            //    }

            //}


            consultantTable.Rows[1].Cells[0].Blocks.Add(
                new Paragraph(document, AddHeading(document, "Name"))
                );

            //Degree
            consultantTable.Rows[1].Cells[0].Blocks.Add(new Paragraph(document, "Degree"));

            //Mobile & Email
            consultantTable.Rows[2].Cells[0].Blocks.Add(new Paragraph(document, Size10(document, "Mobile: " + "Mobile Number")));
            consultantTable.Rows[2].Cells[0].Blocks.Add(new Paragraph(document, Size10(document, "E-mail: " + "Email")));

            return document;
        }
        public static void AddProfile(object cv, List<Inline> items, string docuform, DocumentModel document)
        {
            if (cv != null)
            {
                items.Add(CVDocument.AddHeading(document, "Profile"));
                items.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));
                items.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));
                for (int i = 0; i < 7; i++)
                {
                    if (i % 2 == 0)
                    {
                        items.Add(new Run(document, ""));
                        items.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));
                    }
                    else
                    {
                        items.Add(new Run(document, "For who thoroughly her boy estimating conviction. Removed demands expense account in outward tedious do. Particular way thoroughly unaffected projection favourable mrs can projecting own."));
                        items.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));
                    }

                }

            }
        }
        public static void AddCompetence(object cv, List<Inline> items, string docuform, DocumentModel document, object databasecontext)
        {
            //Create list style.
            ListStyle bulletList = new ListStyle(ListTemplateType.Bullet);

            if (cv != null)
            {

                Paragraph mainpara = new Paragraph(document, new Run(document, "Name" + "'s main competencies are:"));
                document.Sections[0].Blocks.Add(mainpara);

                List<string> dupelist = new List<string>();
                dupelist.Add("fruit");
                dupelist.Add("vegetable");
                List<string> skills = new List<string>();
                skills.Add("apple");
                skills.Add("apple2");
                skills.Add("apple3");
                skills.Add("apple4");
                skills.Add("apple5");
                List<string> skills2 = new List<string>();
                skills2.Add("apple6");
                skills2.Add("apple7");
                skills2.Add("apple8");
                skills2.Add("apple9");
                skills2.Add("apple0");
                foreach (var category in dupelist)
                {
                    //var cats = _context.CompetenceCategoryTranslations.Where(i => i.CompetenceCategoryId == comp.Competence.CompetenceCategoryId).FirstOrDefault();
                    //CompetenceTranslation compTrans = _context.CompetenceTranslations.Where(i => i.CompetenceId == comp.CompetenceId).FirstOrDefault();
                  
                        List<string> allcomps = new List<string>();
                        string catname = dupelist[0] + ": ";

                    if (catname == "fruit")
                    {
                        for (int i = 0; i < skills.Count; i++)
                        {
                            allcomps.Add(skills[i]);
                        }
                    }

                    for (int i = 0; i < skills2.Count; i++)
                    {
                        allcomps.Add(skills2[i]);
                    }

                    var completedString = catname + string.Join(", ", allcomps);
                        items.Add(new Run(document, completedString));
                        document.Sections[0].Blocks.Add(new Paragraph(document, completedString)
                        {
                            ParagraphFormat = new ParagraphFormat() { NoSpaceBetweenParagraphsOfSameStyle = true },
                            ListFormat = new ListFormat() { Style = bulletList }
                        });

                    
                }
            }
        }
        public static Run AddHeading(DocumentModel document, string headingText)
        {
            return new Run(document, headingText)
            {
                CharacterFormat = new CharacterFormat()
                {
                    Size = 16,
                    Bold = true,


                    FontName = "Lacuna Regular"
                }
            };
        }
        public static Paragraph Indent(DocumentModel document, string projectText)
        {
            return new Paragraph(document, projectText)
            {
                ParagraphFormat = new ParagraphFormat
                {
                    LeftIndentation = 25
                }
            };
        }
        public static Run Size10(DocumentModel document, string headingText)
        {
            return new Run(document, headingText)
            {
                CharacterFormat = new CharacterFormat()
                {
                    Size = 10
                }
            };
        }

    }
}