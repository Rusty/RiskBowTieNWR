using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using Attribute = SC.API.ComInterop.Models.Attribute;

namespace RiskBowTieNWR.Helpers
{
    public class RiskModel
    {
        // categories
        private const string _risk = "Risk";
        private const string _ewi = "Early Warning Indicators";
        private const string _cause = "Causes";
        private const string _causeControls = "Cause Controls";
        private const string _causeControlActions = "Cause Control Actions";
        private const string _consequence = "Consequences";
        private const string _consequenceControls = "Consequence Controls";
        private const string _consequenceActions = "Consequence Control Actions";
        private static readonly string [] _categoryNames = {_causeControlActions, _causeControls, _cause, _risk, _consequence, _consequenceControls, _consequenceActions, _ewi};

        private const string _attrVersion = "Version";
        private const string _attrOwner = "Owner.";
        private const string _attrManager = "Manager";
        private const string _attrLinkedControls = "LinkedControls";
        private const string _attrLinkedControlsTypes = "LinkedControlsTypes";
        private const string _attrRationale = "Rationale (Overall)";
        private const string _attrRationaleSafety = "Rationale (Safety)";
        private const string _attrRationalePerformance = "Rationale (Performance)";
        private const string _attrRationaleValue = "Rationale (Value/Finance)";
        private const string _attrRationalePolitical = "Rationale (Political/Reputation)";
        private static readonly string[] _textFields = { _attrVersion, _attrOwner, _attrManager, _attrLinkedControls, _attrLinkedControlsTypes, _attrRationale,
            _attrRationaleSafety, _attrRationalePerformance, _attrRationaleValue, _attrRationalePolitical};

        private const string _attrBaseline = "Base-line";
        private const string _attrRevision = "Revised";
        private const string _attrlastUpdate = "Last Update";
        private static readonly string[] _dateFields = { _attrBaseline, _attrRevision, _attrlastUpdate };

        private const string _attrPercComplete = "% Complete";
        private const string _attrOrder = "SortOrder";
        private const string _attrPrior = "Prior";
        private const string _attrCurrent = "Current";
        private static readonly string[] _numberFields = { _attrPercComplete, _attrOrder, _attrPrior, _attrCurrent };

        private const string _attrControlOpinion = "Control Opinion";
        private const string _attrBasisOfOpinion = "Basis of Opinion";
        private const string _attrPriority = "Priority";
        private const string _attrStatus = "Status";
        private const string _attrClassification = "Classification";
        private const string _attrImpactedArea = "Key Scorecard Area Impacted";
        private const string _attrControlRating = "Control Rating";
        private const string _attrRiskLevel = "Risk Level";
        private const string _attrRiskAppetite = "Risk Appetite";
        private const string _attrRiskAppetiteSafety = "Risk Appetite (Safety)";
        private const string _attrRiskAppetitePerformance = "Risk Appetite (Performance)";
        private const string _attrRiskAppetiteValue = "Risk Appetite (Value/Finance)";
        private const string _attrRiskAppetitePolitical = "Risk Appetite (Political/Reputation)";
        private const string _attrReportingPriority = "Reporting Priority";
        private const string _attrDirectorate = "Directorate";
        private const string _attrGrossRating = "Gross Rating";
        private const string _attrTargetRating = "Target Rating";
        private static readonly string[] _listFields = { _attrControlOpinion, _attrBasisOfOpinion, _attrPriority, _attrStatus, _attrClassification, _attrImpactedArea, _attrControlRating, _attrRiskLevel,
            _attrRiskAppetite, _attrRiskAppetiteSafety, _attrRiskAppetitePerformance, _attrRiskAppetiteValue,_attrRiskAppetitePolitical, _attrReportingPriority, _attrDirectorate,
            _attrGrossRating,  _attrTargetRating  };

        private const string _attrLikelihood = "Likelihood (Overall)";
        private const string _attrImpact = "Impact (Overall)";
        private const string _attrLikelihoodSafety = "Likelihood (Safety)";
        private const string _attrImpactSafety = "Impact (Safety)";
        private const string _attrLikelihoodPerformance = "Likelihood (Performance)";
        private const string _attrImpactPerformance = "Impact (Performace)";
        private const string _attrLikelihoodValue = "Likelihood (Value/Finance)";
        private const string _attrImpactValue = "Impact (Value/Finance)";
        private const string _attrLikelihoodPolitical = "Likelihood (Political/Reputation)";
        private const string _attrImpactPolitical = "Impact (Political/Reputation)";
        private const string _attrGrossImpact = "Gross Impact";
        private const string _attrGrossLkelihood = "Gross Likelihood";
        private const string _attrGrossFinance = "Gross Finance";
        private const string _attrTargetImpact = "Target Impact";
        private const string _attrTargetLkelihood = "Target Likelihood";
        private const string _attrTargetFinance = "Target Finance";
        private static readonly string[] _listRiskFields = { _attrLikelihood, _attrImpact, _attrLikelihoodSafety, _attrImpactSafety, _attrLikelihoodPerformance, _attrImpactPerformance,
            _attrLikelihoodValue, _attrImpactValue, _attrLikelihoodPolitical, _attrImpactPolitical,
             _attrGrossImpact, _attrGrossLkelihood, _attrGrossFinance, _attrTargetImpact, _attrTargetLkelihood, _attrTargetFinance};

        private static readonly string[] _listCauses = { };
        private static readonly string[] _listCausesControls = { _attrOwner, _attrControlOpinion, _attrBasisOfOpinion };
        private static readonly string[] _listCausesActions = { _attrOwner, _attrPriority, _attrBaseline, _attrPercComplete, _attrStatus };
        private static readonly string[] _listConsequenses = { };
        private static readonly string[] _listConsequensesControls = { _attrOwner, _attrControlOpinion, _attrBasisOfOpinion };
        private static readonly string[] _listConsequensesActions = { _attrOwner, _attrPriority, _attrBaseline, _attrPercComplete, _attrStatus };
        private static readonly string[] _listEWI = { };

        private static readonly double[] _widthCauses = { 100 };
        private static readonly double[] _widthCausesControls = { 60, 20, 15, 5 };
        private static readonly double[] _widthCausesActions = { 40, 10, 10, 10, 10, 10, 10 };
        private static readonly double[] _widthConsequenses = { 100 };
        private static readonly double[] _widthConsequensesControls = { 60, 20, 15, 5 };
        private static readonly double[] _widthConsequensesActions = { 40, 10, 10, 10, 10, 10, 10 };
        private static readonly double[] _widthEWI = { 100 };


        private static readonly string[] _riskLabels = {"1-Very Low", "2-Low", "3-Medium", "4-High", "5-Very High"};


        private const string _riskId = "RISK";
        private const string _ewiId = "EWI";
        private const string _causeId = "CAUSE";
        private const string _causeControlsId = "CAUSE_CONTROL";
        private const string _causeControlActionsId = "CAUSE_ACTION";
        private const string _consequenceId = "CONSQ";
        private const string _consequenceControlsId = "CONSQ_CONTROL";
        private const string _consequenceControlActionsId = "CONSQ_ACTION";

        private const string _multipleValues = "Multiple Values";

        public static void EnsureStoryHasRightStructure(Story story, Logger log)
        {
            // make sure all categories we need exist
            foreach (var c in _categoryNames)
            {
                if (story.Category_FindByName(c) == null) // catagory does not exist
                {
                    log.Log($"Adding Category '{c}'");
                    story.Category_AddNew(c);
                }
            }
            // make sure all attributes we need exist
            foreach (var a in _textFields)
            {
                if (story.Attribute_FindByName(a) == null)
                {
                    log.Log($"Adding Text Attribute '{a}'");
                    story.Attribute_Add(a, Attribute.AttributeType.Text);
                }
            }
            foreach (var a in _dateFields)
            {
                if (story.Attribute_FindByName(a) == null)
                {
                    log.Log($"Adding Date Attribute '{a}'");
                    story.Attribute_Add(a, Attribute.AttributeType.Date);
                }
            }
            foreach (var a in _numberFields)
            {
                if (story.Attribute_FindByName(a) == null)
                {
                    log.Log($"Adding Number Attribute '{a}'");
                    story.Attribute_Add(a, Attribute.AttributeType.Numeric);
                }
            }
            foreach (var a in _listFields)
            {
                if (story.Attribute_FindByName(a) == null)
                {
                    log.Log($"Adding List Attribute '{a}'");
                    story.Attribute_Add(a, Attribute.AttributeType.List);
                }
            }
            // reserved list attributes need VL,L,M,H,VH values
            foreach (var a in _listRiskFields)
            {
                if (story.Attribute_FindByName(a) == null)
                {
                    log.Log($"Adding List Attribute '{a}'");
                    var att = story.Attribute_Add(a, Attribute.AttributeType.List);
                    foreach (var l in _riskLabels)
                    {
                        log.Log($"Adding List Label '{l}'");
                        att.Labels_Add(l);
                    }
                }
            }
        }

        public static string LookupControlOpinion(string o)
        {
            // Converts 'E' or 'I' to full word
            switch (o)
            {
                case "E":
                    return "Effective";
                case "I":
                    return "Ineffective";
            }
            return o;// notfound
        }


        public static string LookupRiskLabel(string l)
        {
            // converts 1,2,3,4,5 to 1-Very Low etc.
            switch (l)
            {
                case "1":
                    return _riskLabels[0];
                case "2":
                    return _riskLabels[1];
                case "3":
                    return _riskLabels[2];
                case "4":
                    return _riskLabels[3];
                case "5":
                    return _riskLabels[4];
            }
            return l;// notfound
        }

        public static string GetReportingPriority(int order)
        {
            // makes sure top 5 items are reported as high priority (for filtering)    
            if (order <= 5)
                return "High Priority";
            return "Lower Priority";
        }

        public static string GetShortenedDirectorate(string str)
        {
            str = str.Replace("&", "").Replace("  ", " ");

            var words = str.Split(' ');

            return words.Where(w => w.Length >= 1).Aggregate("", (current, w) => current + w[0]);
        }


        private const string _strRiskRelatedStory = "Risk related story";
        private const string _strRiskControlInstance = "Instance of this control";
        
        public static async void ProcessBowTies(SharpCloudApi sc, string teamId, string portfolioId, string controlId, string temaplateId, Logger log)
        {
            // combine all bowtie stories 
            log.Log($"Reading Portfolio Story");
            await Task.Delay(100);
            var portfolioStory = sc.LoadStory(portfolioId);

            log.Log($"Reading Control Story");
            await Task.Delay(100);
            var controlsStory = sc.LoadStory(controlId);

            // make sure risk count exists
            var _riskCount = "RiskCount";
            var attRiskCountControlStory = controlsStory.Attribute_FindByName(_riskCount) ??
                               controlsStory.Attribute_Add(_riskCount, Attribute.AttributeType.Numeric);
            var attOverallControlRatingControlStory = controlsStory.Attribute_FindByName(_attrControlRating) ??
                   controlsStory.Attribute_Add(_attrControlRating, Attribute.AttributeType.List);
            var attManagedControlStory = controlsStory.Attribute_FindByName(_attrControlOpinion) ??
                   controlsStory.Attribute_Add(_attrControlOpinion, Attribute.AttributeType.List);
            var attRiskLevelControlStory = controlsStory.Attribute_FindByName(_attrRiskLevel) ??
                   controlsStory.Attribute_Add(_attrRiskLevel, Attribute.AttributeType.List);
            var attBasisOfOpinionControlStory = controlsStory.Attribute_FindByName(_attrBasisOfOpinion) ??
                   controlsStory.Attribute_Add(_attrBasisOfOpinion, Attribute.AttributeType.List);
            var attKeyScorecardAreaControlStory = controlsStory.Attribute_FindByName(_attrImpactedArea) ??
                               controlsStory.Attribute_Add(_attrImpactedArea, Attribute.AttributeType.List);


            /* Don't remove
            attOverallControlRating.Labels_Delete(_multipleValues); // always removed
            attManaged.Labels_Delete(_multipleValues); // always removed
            */

            // remove control ratings for existing controls

            foreach (var itm in controlsStory.Items)
            {
                itm.RemoveAttributeValue(attManagedControlStory);
                itm.RemoveAttributeValue(attOverallControlRatingControlStory);
                itm.RemoveAttributeValue(attRiskLevelControlStory);
                itm.RemoveAttributeValue(attBasisOfOpinionControlStory);
                itm.RemoveAttributeValue(attKeyScorecardAreaControlStory);

                itm.SetAttributeValue(attRiskCountControlStory, 0);

                // delete all resources - we will add new ones back in later
                var list = new List<string>();
                foreach (var res in itm.Resources)
                {
                    if (res.Description == _strRiskControlInstance || res.Description == _strRiskRelatedStory)
                        list.Add(res.Id);
                }
                foreach(var id in list)
                    itm.Resource_DeleteById(id);
            }



            foreach (var teamStory in sc.StoriesTeam(teamId))
            {
                if (teamStory.Id != portfolioId && teamStory.Id != temaplateId && teamStory.Id != controlId)// && teamStory.Id == "aad81010-af26-4ba2-954a-420383fb6d1f")
                {
                    log.Log($"Reading from '{teamStory.Name}'");
                    await Task.Delay(100);

                    
                    try
                    {
                        var story = sc.LoadStory(teamStory.Id);
                        var riskItemSource = story.Item_FindByExternalId("RISK");

                        if (riskItemSource != null)
                        {
                            var riskItem = portfolioStory.Item_FindByExternalId(story.Id) ??
                               portfolioStory.Item_AddNew(riskItemSource.Name, false);

                            riskItem.Name = riskItemSource.Name;
                            riskItem.Description = riskItemSource.Description;
                            riskItem.ExternalId = story.Id;

                            var attDirectorate = story.Attribute_FindByName(_attrDirectorate);
                            var riskCategoryName = riskItemSource.GetAttributeValueAsText(attDirectorate);

                            var riskCategory = portfolioStory.Category_FindByName(riskCategoryName) ??
                                               portfolioStory.Category_AddNew(riskCategoryName);

                            riskItem.Category = riskCategory;


                            var res = riskItem.Resource_FindByName(teamStory.Name) ?? riskItem.Resource_AddName(teamStory.Name);
                            res.Description = story.Description;
                            res.Url = new Uri(story.Url);

                            LoadPanelData(riskItem, story, _cause, _listCauses, _widthCauses);
                            LoadPanelData(riskItem, story, _causeControls, _listCausesControls, _widthCausesControls);
                            LoadPanelData(riskItem, story, _causeControlActions, _listCausesActions, _widthCausesActions);
                            LoadPanelData(riskItem, story, _consequence, _listConsequenses, _widthConsequenses);
                            LoadPanelData(riskItem, story, _consequenceControls, _listConsequensesControls, _widthConsequensesControls);
                            LoadPanelData(riskItem, story, _consequenceActions, _listConsequensesActions, _widthConsequensesActions);
                            LoadPanelData(riskItem, story, _ewi, _listEWI, _widthEWI);

                            if (riskItemSource != null)
                                CopyAllAttributeValues(riskItemSource, riskItem);
                            else
                                log.LogError($"Could not find a risk item in {teamStory.Name}");


                            // do the controls
                            var attManagedControlsRiskStory = story.Attribute_FindByName(_attrControlOpinion) ??
                                                     story.Attribute_Add(_attrControlOpinion,
                                                         Attribute.AttributeType.List);
                            var attRiskLevelControlsRiskStory = story.Attribute_FindByName(_attrRiskLevel) ??
                                                       story.Attribute_Add(_attrRiskLevel, Attribute.AttributeType.List);
                            var attOverallControlRatingControlsRiskStory = story.Attribute_FindByName(_attrControlRating) ??
                                                                  story.Attribute_Add(_attrControlRating,
                                                                      Attribute.AttributeType.List);
                            var attBasisOfOpinionRiskStory = story.Attribute_FindByName(_attrBasisOfOpinion) ??
                                   controlsStory.Attribute_Add(_attrBasisOfOpinion, Attribute.AttributeType.List);

                            var attScorecardAreaRiskStory = story.Attribute_FindByName(_attrImpactedArea) ??
                                   controlsStory.Attribute_Add(_attrImpactedArea, Attribute.AttributeType.List);


                            foreach (
                                var itemControlSource in
                                    story.Items.Where(
                                        i =>
                                            (i.Category.Name == _causeControls ||
                                             i.Category.Name == _consequenceControls)))
                            {
                                var itemControlDestination = controlsStory.Item_FindByName(itemControlSource.Name) ??
                                            controlsStory.Item_AddNew(itemControlSource.Name);

                                itemControlDestination.Description = itemControlSource.Description;

                                itemControlDestination.Tag_AddNew(itemControlSource.Category.Name);
                                // add any tags
                                foreach (var t in itemControlSource.Tags)
                                {
                                    itemControlDestination.Tag_AddNew(t.Text);
                                }

                                var resC = itemControlDestination.Resource_FindByUrl(itemControlSource.Url);
                                if (resC == null)
                                    resC = itemControlDestination.Resource_AddName(itemControlSource.Id); // need to be unique
                                resC.Name = "Usage of this control";
                                resC.Description = $"Control used in '{story.Name}'";//"Instance of This Control";
                                resC.Url = new Uri(itemControlSource.Url);

                                var resCC = itemControlSource.Resource_FindByUrl(itemControlDestination.Url);
                                if (resCC == null)
                                    resCC = itemControlSource.Resource_AddName(itemControlDestination.Id); // need to be unique
                                resCC.Name = "Control Library";
                                resCC.Description = "View this control in Control library";
                                resCC.Url = new Uri(itemControlDestination.Url);


                                // risk control
                                AddAttributeButCheckForDiffernce(itemControlSource, attManagedControlsRiskStory, itemControlDestination, attManagedControlStory);
                                // overall risk rating
                                AddAttributeButCheckForDiffernce(riskItemSource, attOverallControlRatingControlsRiskStory, itemControlDestination, attOverallControlRatingControlStory);
                                // overall risk Level
                                AddAttributeButCheckForDiffernce(riskItemSource, attRiskLevelControlsRiskStory, itemControlDestination, attRiskLevelControlStory);
                                // add risk basis of opinion
                                AddAttributeButCheckForDiffernce(itemControlSource, attBasisOfOpinionRiskStory, itemControlDestination, attBasisOfOpinionControlStory);
                                // add risk group (category)
                                AddAttributeButCheckForDiffernce(riskItemSource, attScorecardAreaRiskStory, itemControlDestination, attKeyScorecardAreaControlStory);

                            }
                            story.Save();// save resourc links
                        }
                        else
                        {
                            log.LogError($"Could not find a risk item in {teamStory.Name}");
                        }

                    }
                    catch (Exception e)
                    {
                        log.LogError(e.Message);
                    }
                }
            }

            // process control item resource panels
            foreach (var itm in controlsStory.Items)
            {
                // delete all resources - we will add new ones back in later
                var listR = new List<string>();
                var listC = new List<string>();
                foreach (var res in itm.Resources)
                {
                    if (res.Description == _strRiskControlInstance)
                        listC.Add(res.Id);
                    if (res.Description == _strRiskRelatedStory)
                        listR.Add(res.Id);
                }
            }
 
            log.Log($"Saving {portfolioStory.Name}");
            portfolioStory.Save();
            log.Log($"Saving {controlsStory.Name}");
            controlsStory.Save();
            await Task.Delay(1000);

            log.HideProgress();
        }

        private static void AddAttributeButCheckForDiffernce(Item sourceItem, Attribute sourceAttrib, Item destItem, Attribute destAttrib)
        {
            var test = sourceItem.GetAttributeValueAsText(sourceAttrib);
            if (destItem.GetAttributeIsAssigned(destAttrib))
            {
                // check values are same
                var testDest = destItem.GetAttributeValueAsText(destAttrib);
                if (test != testDest)
                {
                    destItem.SetAttributeValue(destAttrib, _multipleValues);
                    return;
                }

            }
            // else can set
            destItem.SetAttributeValue(destAttrib, test);
        }

        private static void CopyAllAttributeValues(Item itemSource, Item itemDestination)
        {
            foreach (var attrib in itemSource.Story.Attributes)
            {
                if (itemSource.GetAttributeIsAssigned(attrib)) // only copy the value if its assigned
                {
                    var attD = EnsureAttributeExists(attrib, itemDestination.Story);

                    itemDestination.SetAttributeValue(attD, itemSource.GetAttributeValueAsText(attrib));
                }
            }
        }

        private static Attribute EnsureAttributeExists(Attribute attrib, Story destination)
        {
            var attDest = destination.Attribute_FindByName(attrib.Name);
            if (attDest == null)
            {
                attDest = destination.Attribute_Add(attrib.Name, attrib.Type);
                foreach (var lab in attrib.Labels)
                {
                    attDest.Labels_Add(lab.Text, lab.Color);
                }
            }
            return attDest;
        }


        private static void LoadPanelData(Item item, Story story, string category, string[] attributes, double[] widths = null)
        {
            var sortby = story.Attribute_FindByName("SortOrder");
            
            // add attributes so it won't blow up below
            foreach (var attribute in attributes)
            {
                if (story.Attribute_FindByName(attribute) == null)
                    story.Attribute_Add(attribute, Attribute.AttributeType.List);
            }

            var table1 = new HTMLTable(attributes.Length + 1);
            int col = 0;
            table1.SetValue(0, col++, "Name");
            foreach (var attribute in attributes)
            {
                if (widths != null)
                    table1.SetColWidth(col, widths[col]);
                table1.SetValue(0, col++, attribute);
            }

            int row = 1;
            foreach (var item2 in story.Items.OrderBy(i => i.GetAttributeValueAsDouble(sortby)))
            {
                if (item2.Category.Name == category)
                {
                    //Debug.WriteLine($"Category = '{category}'");
                    col = 0;
                    table1.SetValue(row, col++, item2.Name);
                    //Debug.WriteLine($"Item = '{item2.Name}'");
                    foreach (var attribute in attributes)
                    {
                        //Debug.WriteLine($"Attrubute = '{attribute}'");
                        var att = story.Attribute_FindByName(attribute);
                        //Debug.WriteLine($"Attrubute.Name = '{att.Name}'");
                        var text = item2.GetAttributeValueAsTextWithPrefixAndSuffix(att);
                        //Debug.WriteLine($"Text = '{text}'");
                        string color = item2.GetAttributeValueAsColorText(att);
                        table1.SetValue(row, col++, text, color);
                    }
                    row++;
                }
            }

            var panel = item.Panel_FindByTitle(category);
            if (panel == null)
                panel = item.Panel_Add(category, Panel.PanelType.RichText);
            panel.Data = table1.GetHTML;

        }

        public static string GetExcelTemplateStoryID(string XLFilename, Logger log)
        {
            var XL1 = new Application();
            var pathMlstn = XLFilename;
            log.Log($"Opening Excel Doc " + pathMlstn);
            var wbBowTie = XL1.Workbooks.Open(pathMlstn);

            var id = XL1.Sheets[1].Cells(12, 4).Text;

            wbBowTie.Close(false);
            XL1.Quit();
            return id;
        }

        public static void SetExcelTemplateStoryID(string Id, string XLFilename, Logger log)
        {
            var XL1 = new Application();
            var pathMlstn = XLFilename;
            log.Log($"Opening Excel Doc " + pathMlstn);
            var wbBowTie = XL1.Workbooks.Open(pathMlstn);

            XL1.Sheets[1].Cells[12, 4] = Id;

            wbBowTie.Close(true);
            XL1.Quit();
        }


        public static void CreateStoryFromXLTemplate(Story story, string XLFilename, Logger log)
        {
            EnsureStoryHasRightStructure(story, log);

            // get categories
            var catRisk = story.Category_FindByName(_risk);
            var catEWI = story.Category_FindByName(_ewi);
            var catCause = story.Category_FindByName(_cause);
            var catCauseControl = story.Category_FindByName(_causeControls);
            var catCauseAction = story.Category_FindByName(_causeControlActions);
            var catConsequence = story.Category_FindByName(_consequence);
            var catConsequenceControl = story.Category_FindByName(_consequenceControls);
            var catConsequenceAction = story.Category_FindByName(_consequenceActions);

            // get attributes
            var attLikelihood = story.Attribute_FindByName(_attrLikelihood);
            var attImpact = story.Attribute_FindByName(_attrImpact);
            var attRationale = story.Attribute_FindByName(_attrRationale);

            var attLikelihoodSafety = story.Attribute_FindByName(_attrLikelihoodSafety);
            var attImpactSafety = story.Attribute_FindByName(_attrImpactSafety);
            var attAppetiteSafety = story.Attribute_FindByName(_attrRiskAppetiteSafety);
            var attRationaleSafety = story.Attribute_FindByName(_attrRationaleSafety);

            var attLikelihoodPerformance = story.Attribute_FindByName(_attrLikelihoodPerformance);
            var attImpactPerformance = story.Attribute_FindByName(_attrImpactPerformance);
            var attAppetitePerformace = story.Attribute_FindByName(_attrRiskAppetitePerformance);
            var attRationalePerformance = story.Attribute_FindByName(_attrRationalePerformance);

            var attLikelihoodValue = story.Attribute_FindByName(_attrLikelihoodValue);
            var attImpactValue = story.Attribute_FindByName(_attrImpactValue);
            var attAppetiteValue = story.Attribute_FindByName(_attrRiskAppetiteValue);
            var attRationaleValue = story.Attribute_FindByName(_attrRationaleValue);

            var attLikelihoodPolitical = story.Attribute_FindByName(_attrLikelihoodPolitical);
            var attImpactPolitical = story.Attribute_FindByName(_attrImpactPolitical);
            var attAppetitePolitical = story.Attribute_FindByName(_attrRiskAppetitePolitical);
            var attRationalePolitical = story.Attribute_FindByName(_attrRationalePolitical);

            var attGrossImpact = story.Attribute_FindByName(_attrGrossImpact);
            var attGrossLikelihood = story.Attribute_FindByName(_attrGrossLkelihood);
            var attGrossFinance = story.Attribute_FindByName(_attrGrossFinance);
            var attGrossRating = story.Attribute_FindByName(_attrGrossRating);
            var attTargetImpact = story.Attribute_FindByName(_attrTargetImpact);
            var attTargetLikelihood = story.Attribute_FindByName(_attrTargetLkelihood);
            var attTargetFinance = story.Attribute_FindByName(_attrTargetFinance);
            var attTargetRating = story.Attribute_FindByName(_attrTargetRating);

            var attControlRating = story.Attribute_FindByName(_attrControlRating);
            var attImapactedArea = story.Attribute_FindByName(_attrImpactedArea);
            var attOwner = story.Attribute_FindByName(_attrOwner);
            var attBasisOfOpinion = story.Attribute_FindByName(_attrBasisOfOpinion);
            var attLinkedControls = story.Attribute_FindByName(_attrLinkedControls);
            var attLinkedControlsTypes = story.Attribute_FindByName(_attrLinkedControlsTypes);
            var attBaseline = story.Attribute_FindByName(_attrBaseline);
            var attRevision = story.Attribute_FindByName(_attrRevision);
            var attPercComplete = story.Attribute_FindByName(_attrPercComplete);
            var attSortOrder = story.Attribute_FindByName(_attrOrder);
            var attPrior = story.Attribute_FindByName(_attrPrior);
            var attCurrent = story.Attribute_FindByName(_attrCurrent);
            var attControlOpinion = story.Attribute_FindByName(_attrControlOpinion);
            var attPriority = story.Attribute_FindByName(_attrPriority);
            var attStatus = story.Attribute_FindByName(_attrStatus);
            var attClassification = story.Attribute_FindByName(_attrClassification);
            var attVersion = story.Attribute_FindByName(_attrVersion);
            var attLastUpdate = story.Attribute_FindByName(_attrlastUpdate);
            var attManager = story.Attribute_FindByName(_attrManager);
            var attRiskLevel = story.Attribute_FindByName(_attrRiskLevel);
            var attReportingPriority = story.Attribute_FindByName(_attrReportingPriority);
            var attDirectorate = story.Attribute_FindByName(_attrDirectorate);


            try
            {

                var XL1 = new Application();
                var pathMlstn = XLFilename;
                log.Log($"Opening Excel Doc " + pathMlstn);
                var wbBowTie = XL1.Workbooks.Open(pathMlstn);
                var sheet = 1;

                // validate template is correct version
                var version = XL1.Sheets["Version Control"].Cells[1, 26].Text;
                if (version != "SCApproved")
                {
                    log.Log($"Spreadhseet is not in the approved version, missing 'SCApproved' at Z1 in 'Version Control' ");
                    KillProcessByMainWindowHwnd(XL1.Application.Hwnd);
                    return;
                }


                // set the story name
                var level = XL1.Sheets[sheet].Cells(3, 4).Text;
                var directorate = XL1.Sheets[sheet].Cells(4, 4).Text;
                var title = XL1.Sheets[sheet].Cells(5, 4).Text;
                
                story.Name = $"L{level}_{GetShortenedDirectorate(directorate)}_{title}";

                Item risk = story.Item_FindByExternalId(_riskId) ?? story.Item_AddNew(title, false);
                risk.ExternalId = _riskId;
                risk.Description = XL1.Sheets[sheet].Cells(3, 19).Text;
                risk.Category = catRisk;
                SetAttributeWithLogging(log, risk, attClassification, XL1.Sheets[sheet].Cells(2, 4).Text);
                SetAttributeWithLogging(log, risk, attRiskLevel, level);
                SetAttributeWithLogging(log, risk, attDirectorate, directorate);

                SetAttributeWithLogging(log, risk, attOwner, XL1.Sheets[sheet].Cells(6, 4).Text);
                SetAttributeWithLogging(log, risk, attManager, XL1.Sheets[sheet].Cells(7, 4).Text);
                SetAttributeWithLogging(log, risk, attImapactedArea, LookupRiskLabel(XL1.Sheets[sheet].Cells(8, 4).Text));
                SetAttributeWithLogging(log, risk, attControlRating, XL1.Sheets[sheet].Cells(9, 4).Text);
                SetAttributeWithLogging(log, risk, attVersion, XL1.Sheets[sheet].Cells(10, 4).Text);
                SetAttributeWithLogging(log, risk, attLastUpdate, XL1.Sheets[sheet].Cells(11, 4).Text);

                // gross
                SetAttributeWithLogging(log, risk, attGrossImpact, LookupRiskLabel(XL1.Sheets[sheet].Cells(15, 4).Text));
                SetAttributeWithLogging(log, risk, attGrossLikelihood, LookupRiskLabel(XL1.Sheets[sheet].Cells(15, 7).Text));
                SetAttributeWithLogging(log, risk, attGrossFinance, LookupRiskLabel(XL1.Sheets[sheet].Cells(15, 10).Text));
                SetAttributeWithLogging(log, risk, attGrossRating, XL1.Sheets[sheet].Cells(15, 12).Text);
                // target
                SetAttributeWithLogging(log, risk, attTargetImpact, LookupRiskLabel(XL1.Sheets[sheet].Cells(16, 4).Text));
                SetAttributeWithLogging(log, risk, attTargetLikelihood, LookupRiskLabel(XL1.Sheets[sheet].Cells(16, 7).Text));
                SetAttributeWithLogging(log, risk, attTargetFinance, LookupRiskLabel(XL1.Sheets[sheet].Cells(16, 10).Text));
                SetAttributeWithLogging(log, risk, attTargetRating, XL1.Sheets[sheet].Cells(16, 12).Text);

                SetAttributeWithLogging(log, risk, attLikelihoodSafety, LookupRiskLabel(XL1.Sheets[sheet].Cells(20, 37).Text));
                SetAttributeWithLogging(log, risk, attImpactSafety, LookupRiskLabel(XL1.Sheets[sheet].Cells(20, 35).Text));
                SetAttributeWithLogging(log, risk, attAppetiteSafety, XL1.Sheets[sheet].Cells(22, 35).Text);
                SetAttributeWithLogging(log, risk, attRationaleSafety, XL1.Sheets[sheet].Cells(20, 19).Text);

                SetAttributeWithLogging(log, risk, attLikelihoodSafety, LookupRiskLabel(XL1.Sheets[sheet].Cells(20, 37).Text));
                SetAttributeWithLogging(log, risk, attImpactSafety, LookupRiskLabel(XL1.Sheets[sheet].Cells(20, 35).Text));
                SetAttributeWithLogging(log, risk, attAppetiteSafety, XL1.Sheets[sheet].Cells(22, 35).Text);
                SetAttributeWithLogging(log, risk, attRationaleSafety, XL1.Sheets[sheet].Cells(20, 19).Text);

                SetAttributeWithLogging(log, risk, attLikelihoodPerformance, LookupRiskLabel(XL1.Sheets[sheet].Cells(25, 37).Text));
                SetAttributeWithLogging(log, risk, attImpactPerformance, LookupRiskLabel(XL1.Sheets[sheet].Cells(25, 35).Text));
                SetAttributeWithLogging(log, risk, attAppetitePerformace, XL1.Sheets[sheet].Cells(27, 35).Text);
                SetAttributeWithLogging(log, risk, attRationalePerformance, XL1.Sheets[sheet].Cells(25, 19).Text);

                SetAttributeWithLogging(log, risk, attLikelihoodValue, LookupRiskLabel(XL1.Sheets[sheet].Cells(30, 37).Text));
                SetAttributeWithLogging(log, risk, attImpactValue, LookupRiskLabel(XL1.Sheets[sheet].Cells(30, 35).Text));
                SetAttributeWithLogging(log, risk, attAppetiteValue, XL1.Sheets[sheet].Cells(32, 35).Text);
                SetAttributeWithLogging(log, risk, attRationaleValue, XL1.Sheets[sheet].Cells(30, 19).Text);

                SetAttributeWithLogging(log, risk, attLikelihoodPolitical, LookupRiskLabel(XL1.Sheets[sheet].Cells(35, 37).Text));
                SetAttributeWithLogging(log, risk, attImpactPolitical, LookupRiskLabel(XL1.Sheets[sheet].Cells(35, 35).Text));
                SetAttributeWithLogging(log, risk, attAppetitePolitical, XL1.Sheets[sheet].Cells(37, 35).Text);
                SetAttributeWithLogging(log, risk, attRationalePolitical, XL1.Sheets[sheet].Cells(35, 19).Text);

                SetAttributeWithLogging(log, risk, attLikelihood, LookupRiskLabel(XL1.Sheets[sheet].Cells(2, 79).Text)); //TODO
                SetAttributeWithLogging(log, risk, attImpact, LookupRiskLabel(XL1.Sheets[sheet].Cells(1, 79).Text)); //TODO
                SetAttributeWithLogging(log, risk, attRationale, XL1.Sheets[sheet].Cells(40, 19).Text);

                SetAttributeWithLogging(log, risk, attReportingPriority, GetReportingPriority(0));

                Item item;
                string extId;
                int order;
                int counterEWI = 1;
                string name;
                string desc;
                string text;
                // data can be in teh same place on 3 sheets (continuation sheets)
                for (sheet = 1; sheet <= 3; sheet++)
                {
                    // cause
                    for (int row = 23; row < 43; row += 2)
                    {
                        text = XL1.Sheets[sheet].Cells(row, 3).Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            GetItemNameAndDescription(text, out name, out desc);
                            order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                            extId = _causeId + $"{order:D2}";
                            item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                            item.ExternalId = extId;
                            item.Name = name;
                            item.Description = desc;
                            item.Category = catCause;
                            SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 16).Text);

                            SetAttributeWithLogging(log, item, attSortOrder, order);
                            SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));

                        }
                    }
                    // cause-control
                    for (int row = 48; row < 59; row++)
                    {
                        text = XL1.Sheets[sheet].Cells(row, 3).Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            GetItemNameAndDescription(text, out name, out desc);
                            order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                            extId = _causeControlsId + $"{order:D2}";

                            item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                            item.ExternalId = extId;
                            item.Name = name;
                            item.Description = desc;
                            item.Category = catCauseControl;
                            SetAttributeWithLogging(log, item, attOwner, XL1.Sheets[sheet].Cells(row, 9).Text.Trim());
                            SetAttributeWithLogging(log, item, attControlOpinion,
                                LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 10).Text.Trim()));
                            SetAttributeWithLogging(log, item, attBasisOfOpinion,
                                XL1.Sheets[sheet].Cells(row, 11).Text.Trim());

                            SetAttributeWithLogging(log, item, attSortOrder, order);
                            SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));

                            item.Relationship_AddItem(risk, "", Relationship.RelationshipDirection.AtoB);
                        }
                    }
                    // cause-action
                    for (int row = 48; row < 59; row++)
                    {
                        text = XL1.Sheets[sheet].Cells(row, 14).Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            GetItemNameAndDescription(text, out name, out desc);
                            order = GetInt(XL1.Sheets[sheet].Cells(row, 13).Text.Trim());
                            extId = _causeControlActionsId + $"{order:D2}";

                            item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                            item.ExternalId = extId;
                            item.Name = name;
                            item.Description = desc;
                            item.Category = catCauseAction;

                            SetAttributeWithLogging(log, item, attOwner, XL1.Sheets[sheet].Cells(row, 21).Text.Trim());
                            SetAttributeWithLogging(log, item, attPriority, XL1.Sheets[sheet].Cells(row, 22).Text.Trim());
                            SetAttributeWithLogging(log, item, attBaseline, XL1.Sheets[sheet].Cells(row, 23).Text.Trim());
                            SetAttributeWithLogging(log, item, attRevision, XL1.Sheets[sheet].Cells(row, 24).Text.Trim());
                            SetAttributeWithLogging(log, item, attPercComplete,
                                XL1.Sheets[sheet].Cells(row, 25).Text.Trim().Replace("%", ""));
                            SetAttributeWithLogging(log, item, attStatus, XL1.Sheets[sheet].Cells(row, 26).Text.Trim());

                            SetAttributeWithLogging(log, item, attSortOrder, order);
                            SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));
                        }
                    }

                    // consequences
                    for (int row = 23; row < 43; row += 2)
                    {
                        text = XL1.Sheets[sheet].Cells(row, 41).Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            GetItemNameAndDescription(text, out name, out desc);
                            order = GetInt(XL1.Sheets[sheet].Cells(row, 40).Text.Trim());
                            extId = _consequenceId + $"{order:D2}";

                            item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                            item.ExternalId = extId;
                            item.Name = name;
                            item.Description = desc;
                            item.Category = catConsequence;
                            SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 56).Text);

                            SetAttributeWithLogging(log, item, attSortOrder, order);
                            SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));

                        }
                    }
                    // consequence-control
                    for (int row = 48; row < 59; row++)
                    {
                        text = XL1.Sheets[sheet].Cells(row, 33).Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            GetItemNameAndDescription(text, out name, out desc);
                            order = GetInt(XL1.Sheets[sheet].Cells(row, 32).Text.Trim());
                            extId = _consequenceControlsId + $"{order:D2}";

                            item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                            item.ExternalId = extId;
                            item.Name = name;
                            item.Description = desc;
                            item.Category = catConsequenceControl;
                            SetAttributeWithLogging(log, item, attOwner, XL1.Sheets[sheet].Cells(row, 39).Text.Trim());
                            SetAttributeWithLogging(log, item, attControlOpinion,
                                LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 40).Text.Trim()));
                            SetAttributeWithLogging(log, item, attBasisOfOpinion,
                                XL1.Sheets[sheet].Cells(row, 41).Text.Trim());

                            SetAttributeWithLogging(log, item, attSortOrder, order);
                            SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));

                            item.Relationship_AddItem(risk, "", Relationship.RelationshipDirection.BtoA);
                        }
                    }
                    // consequence-action
                    for (int row = 48; row < 59; row++)
                    {
                        text = XL1.Sheets[sheet].Cells(row, 45).Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            GetItemNameAndDescription(text, out name, out desc);
                            order = GetInt(XL1.Sheets[sheet].Cells(row, 32).Text.Trim());
                            extId = _consequenceControlActionsId + $"{order:D2}";

                            item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                            item.ExternalId = extId;
                            item.Name = name;
                            item.Description = desc;
                            item.Category = catConsequenceAction;

                            SetAttributeWithLogging(log, item, attOwner, XL1.Sheets[sheet].Cells(row, 53).Text.Trim());
                            SetAttributeWithLogging(log, item, attPriority, XL1.Sheets[sheet].Cells(row, 54).Text.Trim());
                            SetAttributeWithLogging(log, item, attBaseline, XL1.Sheets[sheet].Cells(row, 55).Text.Trim());
                            SetAttributeWithLogging(log, item, attRevision, XL1.Sheets[sheet].Cells(row, 56).Text.Trim());
                            SetAttributeWithLogging(log, item, attPercComplete,
                                XL1.Sheets[sheet].Cells(row, 57).Text.Trim().Replace("%", ""));
                            SetAttributeWithLogging(log, item, attStatus, XL1.Sheets[sheet].Cells(row, 58).Text.Trim());

                            SetAttributeWithLogging(log, item, attSortOrder, order);
                            SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));
                        }
                    }

                    for (int row = 9; row <= 17; row += 2)
                    {
                        text = XL1.Sheets[sheet].Cells(row, 22).Text;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            GetItemNameAndDescription(text, out name, out desc);
                            order = counterEWI++;
                            extId = _ewiId + $"{order:D2}";

                            item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                            item.ExternalId = extId;
                            item.Name = name;
                            item.Description = desc;
                            item.Category = catEWI;

                            SetAttributeWithLogging(log, item, attLinkedControlsTypes,
                                XL1.Sheets[sheet].Cells(row, 19).Text);
                            SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 21).Text);

                            SetAttributeWithLogging(log, item, attPrior, XL1.Sheets[sheet].Cells(row, 35).Text);
                            SetAttributeWithLogging(log, item, attCurrent, XL1.Sheets[sheet].Cells(row, 37).Text);

                            SetAttributeWithLogging(log, item, attSortOrder, order);
                            SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));

                            var link = XL1.Sheets[sheet].Cells(row, 19).Text;

                        }
                    }
                }

                // process the relationships for the Causes & Conseqences
                // the template only allows for a max of 30 of each
                for (int c = 1; c <= 30; c++)
                {
                    // Causes
                    extId = _causeId + $"{c:D2}";
                    item = story.Item_FindByExternalId(extId);
                    if (item != null)
                    {
                        var rels = item.GetAttributeValueAsText(attLinkedControls);
                        foreach (var r in rels.Split(','))
                        {
                            var i = GetInt(r);
                            var ex = _causeControlsId + $"{i:D2}";
                            var itm = story.Item_FindByExternalId(ex);
                            if (itm != null)
                            {
                                item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.AtoB);
                            }
                        }
                    }

                    // Consequences
                    extId = _consequenceId + $"{c:D2}";
                    item = story.Item_FindByExternalId(extId);
                    if (item != null)
                    {
                        var rels = item.GetAttributeValueAsText(attLinkedControls);
                        foreach (var r in rels.Split(','))
                        {
                            var i = GetInt(r);
                            var ex = _consequenceControlsId + $"{i:D2}";
                            var itm = story.Item_FindByExternalId(ex);
                            if (itm != null)
                            {
                                item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.BtoA);
                            }
                        }
                    }

                    // Consequences
                    extId = _consequenceId + $"{c:D2}";
                    item = story.Item_FindByExternalId(extId);
                    if (item != null)
                    {
                        var rels = item.GetAttributeValueAsText(attLinkedControls);
                        foreach (var r in rels.Split(','))
                        {
                            var i = GetInt(r);
                            var ex = _consequenceControlsId + $"{i:D2}";
                            var itm = story.Item_FindByExternalId(ex);
                            if (itm != null)
                            {
                                item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.BtoA);
                            }
                        }
                    }

                    // Early Warning Indicators
                    extId = _ewiId + $"{c:D2}";
                    item = story.Item_FindByExternalId(extId);
                    if (item != null)
                    {
                        var relType = item.GetAttributeValueAsText(attLinkedControlsTypes);
                        var rels = item.GetAttributeValueAsText(attLinkedControls);
                        foreach (var r in rels.Split(','))
                        {
                            var i = GetInt(r);
                            var ex = _causeControlsId + $"{i:D2}";
                            if (relType == "Conseq.")
                                ex = _consequenceControlsId + $"{i:D2}";

                            var itm = story.Item_FindByExternalId(ex);
                            if (itm != null)
                            {
                                item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.AtoB);
                            }
                        }
                    }

                }

                GC.Collect();
                GC.WaitForFullGCComplete();
                wbBowTie.Close(false);
                Marshal.ReleaseComObject(wbBowTie);

                KillProcessByMainWindowHwnd(XL1.Application.Hwnd);
 
                GC.Collect();
                GC.WaitForFullGCComplete();
                GC.Collect();
                GC.WaitForFullGCComplete();
            }
            catch ( Exception ex)
            {
                log.LogError(ex.Message);
            }
            
        }

        private static void GetItemNameAndDescription(string text, out string name, out string desc)
        {
            var split = text.Split(new string[] { ";", "\r\n", "\r", "\n"}, 2, StringSplitOptions.RemoveEmptyEntries);
            if (split.Count() == 2)
            {
                name = split[0].Substring(0, Math.Min(200, split[0].Length)).Trim();
                desc = split[1].Substring(0, Math.Min(1000, split[1].Length)).Trim();
            }
            else
            {
                name = text.Substring(0, Math.Min(200, text.Length)).Trim();
                desc = text.Substring(0, Math.Min(1000, text.Length)).Trim();
            }
        }
    
        
        private static void SetAttributeWithLogging(Logger log, Item item, Attribute att, int value)
        {
            SetAttributeWithLogging(log, item, att, $"{value:D2}");
        }

        private static void SetAttributeWithLogging(Logger log, Item item, Attribute att, string value)
        {
            try
            {
                Debug.WriteLine($"{att.Name} {value}");
                item.SetAttributeValue(att, value);
                Debug.WriteLine($"{att.Name} {item.GetAttributeValueAsText(att)}");
            }
            catch (Exception e)
            {
                log.Log($"Error: {e.Message}");
            }
            
        }

        private static int GetInt(string str)
        {
            int ret;
            if (Int32.TryParse(str, out ret))
                return ret;
            return 0;
        }

        public static void MigrateSpreadsheet(string origFile, string newTemplate, string newFile, Logger log)
        {
            if (origFile == newTemplate)
            {
                log.Log($"Skipping, files are the same");
                return;
            }

            var XLS = new Application();
            var pathSource = origFile;
            log.Log($"Opening Excel Doc " + pathSource);

            var wbSource = XLS.Workbooks.Open(pathSource);
            var pathTemplate = newTemplate;
            log.Log($"Opening Excel Doc " + pathTemplate);

            var XLD = new Application();
            var wbTemplate = XLD.Workbooks.Open(pathTemplate);

            // validate template is correct version
            var version = XLD.Sheets["Version Control"].Cells[1, 26].Text;
            if (version != "SCApproved")
            {
                log.Log($"Template in not the approved version, missing 'SCApproved' at Z1 in 'Version Control' ");
                KillProcessByMainWindowHwnd(XLS.Application.Hwnd);
                KillProcessByMainWindowHwnd(XLD.Application.Hwnd);
                return;
            }


            CopyValues(1, XLS, 2, 4, XLD, 2, 4); // Classification 
            CopyValues(1, XLS, 3, 16, XLD, 3, 4); // Risk Level 
            //CopyValues(1, XLS, 4, 4, XLD, 4, 4); // Category 
            CopyValues(1, XLS, 3, 4, XLD, 5, 4); // Risk Title
            CopyValues(1, XLS, 6, 4, XLD, 6, 4); // Risk Owner 
            CopyValues(1, XLS, 7, 4, XLD, 7, 4); // Risk Manager
            for (int r = 10; r <= 15; r++)
            {
                if (XLS.Sheets[1].Cells[r, 7].Text == "a")
                {
                    string sca = XLS.Sheets[1].Cells[r, 1].Text;
                    sca = sca.Replace("Performance", "Business Performance").Replace("Finance & Investment", "Financial Control").Replace("Asset Management", "Asset Stewardship").Replace("Satisfaction / Reputation", "Customer & Stakeholder Relationships").Replace("Additional Measures", "");
                    XLD.Sheets[1].Cells[8, 4] = sca;
                    //CopyValues(1, XLS, r, 1, XLD, 8, 4); // Scorecard Area 
                }
            }

            CopyValues(1, XLS, 16, 8, XLD, 9, 4); // Control Rating 
            CopyValues(1, XLS, 4, 4, XLD, 10, 4); // Version  
            CopyValues(1, XLS, 5, 4, XLD, 11, 4); // Last Update  
            CopyValues(1, XLS, 3, 19, XLD); // risk description

            // safety
            CopyValues(1, XLS, 20, 19, XLD);
            CopyValues(1, XLS, 20, 35, XLD);
            CopyValues(1, XLS, 20, 37, XLD);
            // perf
            CopyValues(1, XLS, 25, 19, XLD);
            CopyValues(1, XLS, 25, 35, XLD);
            CopyValues(1, XLS, 25, 37, XLD);
            // value
            CopyValues(1, XLS, 30, 19, XLD);
            CopyValues(1, XLS, 30, 35, XLD);
            CopyValues(1, XLS, 30, 37, XLD);
            // political
            CopyValues(1, XLS, 35, 19, XLD);
            CopyValues(1, XLS, 35, 35, XLD);
            CopyValues(1, XLS, 35, 37, XLD);
            // overall
            CopyValues(1, XLS, 40, 19, XLD);


            for (int sheet =1; sheet<=3; sheet++)
            {
                // EWI
                for (int row = 9; row <= 17; row += 2)
                {
                    CopyValues(sheet, XLS, row, 19, XLD);
                    CopyValues(sheet, XLS, row, 21, XLD);
                    CopyValues(sheet, XLS, row, 22, XLD);
//                    CopyValues(sheet, XLS, row, 35, XLD); // do not copy prior
                    CopyValues(sheet, XLS, row, 37, XLD, row, 35); // move current column
                }

                for (int row = 23; row <= 41; row += 2)
                {
                    // causes
                    CopyValues(sheet, XLS, row, 3, XLD);
                    CopyValues(sheet, XLS, row, 16, XLD);
                    //conseqeunces
                    CopyValues(sheet, XLS, row, 41, XLD);
                    CopyValues(sheet, XLS, row, 56, XLD);
                }

                for (int row = 48; row <= 59; row += 1)
                {
                    // cause controls
                    CopyValues(sheet, XLS, row, 3, XLD);
                    CopyValues(sheet, XLS, row, 9, XLD);
                    CopyValues(sheet, XLS, row, 10, XLD);
                    CopyValues(sheet, XLS, row, 11, XLD);

                    // cause action
                    CopyValues(sheet, XLS, row, 14, XLD);
                    CopyValues(sheet, XLS, row, 21, XLD);
                    CopyValues(sheet, XLS, row, 22, XLD);
                    CopyValues(sheet, XLS, row, 23, XLD);
                    CopyValues(sheet, XLS, row, 24, XLD);
                    CopyValues(sheet, XLS, row, 25, XLD);
                    CopyValues(sheet, XLS, row, 26, XLD);

                    // consequence controls
                    CopyValues(sheet, XLS, row, 3 + 30, XLD);
                    CopyValues(sheet, XLS, row, 9 + 30, XLD);
                    CopyValues(sheet, XLS, row, 10 + 30, XLD);
                    CopyValues(sheet, XLS, row, 11 + 30, XLD);

                    // conseuence actions
                    CopyValues(sheet, XLS, row, 14 + 31, XLD);
                    CopyValues(sheet, XLS, row, 21 + 31, XLD);
                    CopyValues(sheet, XLS, row, 22 + 31, XLD);
                    CopyValues(sheet, XLS, row, 23 + 31, XLD);
                    CopyValues(sheet, XLS, row, 24 + 31, XLD);
                    CopyValues(sheet, XLS, row, 25 + 31, XLD);
                    CopyValues(sheet, XLS, row, 26 + 31, XLD);
                }
            }

            GC.Collect();
            GC.WaitForFullGCComplete();
            wbSource.Close(false);
            Marshal.ReleaseComObject(wbSource);

            wbTemplate.SaveAs(newFile.Replace("xlsx", "xlsm"), Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);

            wbTemplate.Close(true);
            Marshal.ReleaseComObject(wbTemplate);

            KillProcessByMainWindowHwnd(XLS.Application.Hwnd);
            KillProcessByMainWindowHwnd(XLD.Application.Hwnd);

            //XLS.Quit();
            //Marshal.ReleaseComObject(XLS);
            //XLD.Quit();
            //Marshal.ReleaseComObject(XLD);

            GC.Collect();
            GC.WaitForFullGCComplete();
            GC.Collect();
            GC.WaitForFullGCComplete();

        }

        private static async void CopyValues(int sheet, Application wbS, int row, int col, Application wbD)
        {
            CopyValues(sheet, wbS, row, col, wbD, row, col);
        }

        private static async void CopyValues(int sheet, Application wbS, int rowS, int colS, Application wbD, int rowD, int colD)
        {
            wbD.Sheets[sheet].Cells[rowD, colD] = wbS.Sheets[sheet].Cells[rowS, colS];
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static void KillProcessByMainWindowHwnd(int hWnd)
        {
            uint processID;
            GetWindowThreadProcessId((IntPtr)hWnd, out processID);
            if (processID == 0)
                throw new ArgumentException("Process has not been found by the given main window handle.", "hWnd");
            Process.GetProcessById((int)processID).Kill();
        }
    }
}
