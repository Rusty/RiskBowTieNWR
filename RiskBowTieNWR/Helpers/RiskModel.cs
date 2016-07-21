using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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
        private const string _attrBasisOfOpinion = "Basis of Opinion";
        private const string _attrLinkedControls = "LinkedControls";
        private const string _attrLinkedControlsTypes = "LinkedControlsTypes";
        private const string _attrRationale = "Rationale (Overall)";
        private const string _attrRationaleSafety = "Rationale (Safety)";
        private const string _attrRationalePerformance = "Rationale (Performance)";
        private const string _attrRationaleValue = "Rationale (Value/Finance)";
        private const string _attrRationalePolitical = "Rationale (Political/Reputation)";
        private static readonly string[] _textFields = { _attrVersion, _attrOwner, _attrManager, _attrBasisOfOpinion, _attrLinkedControls, _attrLinkedControlsTypes, _attrRationale,
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
        private static readonly string[] _listFields = { _attrControlOpinion, _attrPriority, _attrStatus, _attrClassification, _attrImpactedArea, _attrControlRating, _attrRiskLevel, _attrRiskAppetite, _attrRiskAppetiteSafety, _attrRiskAppetitePerformance, _attrRiskAppetiteValue,_attrRiskAppetitePolitical };

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
        private static readonly string[] _listRiskFields = { _attrLikelihood, _attrImpact, _attrLikelihoodSafety, _attrImpactSafety, _attrLikelihoodPerformance, _attrImpactPerformance, _attrLikelihoodValue, _attrImpactValue, _attrLikelihoodPolitical, _attrImpactPolitical };

        private static readonly string[] _listCauses = { };
        private static readonly string[] _listCausesControls = { _attrOwner, _attrControlOpinion, _attrBasisOfOpinion };
        private static readonly string[] _listCausesActions = { _attrOwner, _attrPriority, _attrBaseline, _attrPercComplete, _attrStatus };
        private static readonly string[] _listConsequenses = { };
        private static readonly string[] _listConsequensesControls = { _attrOwner, _attrControlOpinion, _attrBasisOfOpinion };
        private static readonly string[] _listConsequensesActions = { _attrOwner, _attrPriority, _attrBaseline, _attrPercComplete, _attrStatus };
        private static readonly string[] _listEWI = { };

        private static readonly double[] _widthCauses = { 1 };
        private static readonly double[] _widthCausesControls = { 0.2, 0.4, 1 };
        private static readonly double[] _widthCausesActions = { 0.3, 0.3, 0.3, 0.3, 0.3, 0.3 };
        private static readonly double[] _widthConsequenses = { 1 };
        private static readonly double[] _widthConsequensesControls = { 0.3, 0.3, 0.3, 0.3 };
        private static readonly double[] _widthConsequensesActions = { 0.3, 0.3, 0.3, 0.3, 0.3, 0.3 };
        private static readonly double[] _widthEWI = { };


        private static readonly string[] _riskLabels = {"1-Very Low", "2-Low", "3-Medium", "4-High", "5-Very High"};


        private const string _riskId = "RISK";
        private const string _ewiId = "EWI";
        private const string _causeId = "CAUSE";
        private const string _causeControlsId = "CAUSE_CONTROL";
        private const string _causeControlActionsId = "CAUSE_ACTION";
        private const string _consequenceId = "CONSQ";
        private const string _consequenceControlsId = "CONSQ_CONTROL";
        private const string _consequenceControlActionsId = "CONSQ_ACTION";

        private static void EnsureStoryHasRightStructure(Story story, Logger log)
        {
            foreach (var c in _categoryNames)
            {
                if (story.Category_FindByName(c) == null) // catagory does not exist
                    story.Category_AddNew(c);
            }

            foreach (var a in _textFields)
            {
                if (story.Attribute_FindByName(a) == null)
                    story.Attribute_Add(a, Attribute.AttributeType.Text);
            }
            foreach (var a in _dateFields)
            {
                if (story.Attribute_FindByName(a) == null)
                    story.Attribute_Add(a, Attribute.AttributeType.Date);
            }
            foreach (var a in _numberFields)
            {
                if (story.Attribute_FindByName(a) == null)
                    story.Attribute_Add(a, Attribute.AttributeType.Numeric);
            }
            foreach (var a in _listFields)
            {
                if (story.Attribute_FindByName(a) == null)
                    story.Attribute_Add(a, Attribute.AttributeType.List);
            }
            foreach (var a in _listRiskFields)
            {
                if (story.Attribute_FindByName(a) == null)
                {
                    var att = story.Attribute_Add(a, Attribute.AttributeType.List);
                    foreach (var l in _riskLabels)
                        att.Labels_Add(l);
                }
            }
        }

        public static string LookupControlOpinion(string o)
        {
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

        public static async void ProcessBowTies(SharpCloudApi sc, string teamId, string portfolioId, string controlId, string temaplateId, Logger log)
        {
            var portfolioStory = sc.LoadStory(portfolioId);
            var controlsStory = sc.LoadStory(controlId);

            foreach (var teamStory in sc.StoriesTeam(teamId))
            {
                if (teamStory.Id != portfolioId && teamStory.Id != temaplateId && teamStory.Id != controlId)// && teamStory.Id == "aad81010-af26-4ba2-954a-420383fb6d1f")
                {
                    log.Log($"Reading from '{teamStory.Name}'");
                    await Task.Delay(100);

                    var riskItem = portfolioStory.Item_FindByName(teamStory.Name) ??
                                   portfolioStory.Item_AddNew(teamStory.Name, false);


                    try
                    {
                        var story = sc.LoadStory(teamStory.Id);
                        var riskItemSource = story.Item_FindByExternalId("RISK");

                        var res = riskItem.Resource_FindByName("Risk Detail");
                        if (res == null)
                            res = riskItem.Resource_AddName("Risk Detail");
                        res.Description = story.Name;
                        res.Url = new Uri(story.Url);

                        LoadPanelData(riskItem, story, _cause, _listCauses);
                        LoadPanelData(riskItem, story, _causeControls, _listCausesControls);
                        LoadPanelData(riskItem, story, _causeControlActions, _listCausesActions);
                        LoadPanelData(riskItem, story, _consequence, _listConsequenses);
                        LoadPanelData(riskItem, story, _consequenceControls, _listConsequensesControls);
                        LoadPanelData(riskItem, story, _consequenceActions, _listConsequensesActions);
                        LoadPanelData(riskItem, story, _ewi, _listEWI);

                        if (riskItemSource != null)
                            CopyAttributeVAlues(riskItemSource, riskItem);
                        else
                            log.LogError($"Could not find a risk item in {teamStory.Name}");

                        foreach (var item1 in story.Items)
                        {
                            if (item1.Category.Name == _causeControls || item1.Category.Name == _consequenceControls)
                            {
                                var cItem = controlsStory.Item_FindByName(item1.Name);
                                if (cItem == null)
                                    cItem = controlsStory.Item_AddNew(item1.Name);

                                cItem.Tag_AddNew(item1.Category.Name);

                                var resC = cItem.Resource_FindByName(story.Name);
                                if (resC == null)
                                    resC = cItem.Resource_AddName(story.Name);
                                resC.Description = "Control used in this risk";
                                resC.Url = new Uri(story.Url);
                            }
                        }


                    }
                    catch (Exception e)
                    {
                        log.LogError(e.Message);
                    }
                }
            }

            log.Log($"Saving {portfolioStory.Name}");
            portfolioStory.Save();
            log.Log($"Saving {controlsStory.Name}");
            controlsStory.Save();
            await Task.Delay(1000);

            log.HideProgress();

        }

        private static void CopyAttributeVAlues(Item itemSource, Item itemDestination)
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




            var XL1 = new Application();
            var pathMlstn = XLFilename;
            log.Log($"Opening Excel Doc " + pathMlstn);
            var wbBowTie = XL1.Workbooks.Open(pathMlstn);
            var sheet = 1;

            Item risk = story.Item_FindByExternalId(_riskId) ?? story.Item_AddNew(XL1.Sheets[sheet].Cells(3, 4).Text, false);
            risk.ExternalId = _riskId;
            risk.Description = XL1.Sheets[sheet].Cells(3, 19).Text;
            risk.Category = catRisk;
            SetAttributeWithLogging(log, risk, attClassification, LookupRiskLabel(XL1.Sheets[sheet].Cells(2, 4).Text));
            SetAttributeWithLogging(log, risk, attVersion, LookupRiskLabel(XL1.Sheets[sheet].Cells(4, 4).Text));
            SetAttributeWithLogging(log, risk, attLastUpdate, LookupRiskLabel(XL1.Sheets[sheet].Cells(4, 5).Text));
            SetAttributeWithLogging(log, risk, attOwner, LookupRiskLabel(XL1.Sheets[sheet].Cells(4, 6).Text));
            SetAttributeWithLogging(log, risk, attManager, LookupRiskLabel(XL1.Sheets[sheet].Cells(4, 7).Text));

            for (int row = 10; row <= 15; row++)
            {
                if (!string.IsNullOrEmpty(XL1.Sheets[sheet].Cells(row, 7).Text))
                    SetAttributeWithLogging(log, risk, attImapactedArea, LookupRiskLabel(XL1.Sheets[sheet].Cells(row, 1).Text));
            }


            SetAttributeWithLogging(log, risk, attControlRating, LookupRiskLabel(XL1.Sheets[sheet].Cells(16, 8).Text));
            SetAttributeWithLogging(log, risk, attRiskLevel, LookupRiskLabel(XL1.Sheets[sheet].Cells(3, 16).Text));

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

            SetAttributeWithLogging(log, risk, attLikelihood, LookupRiskLabel("3")); //TODO
            SetAttributeWithLogging(log, risk, attImpact, LookupRiskLabel("3")); //TODO
            SetAttributeWithLogging(log, risk, attRationale, XL1.Sheets[sheet].Cells(40, 19).Text);

            Item item;
            string extId;
            int order;
            int counterEWI = 1;
            // data can be in teh same place on 3 sheets (continuation sheets)
            for (sheet = 1; sheet <= 3; sheet++)
            {
                // cause
                for (int row = 23; row < 43; row += 2)
                {
                    string name = XL1.Sheets[sheet].Cells(row, 3).Text;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                        extId = _causeId + $"{order:D2}";
                        item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name.Substring(0, Math.Min(50, name.Length)), false);
                        item.ExternalId = extId;
                        item.Description = name;
                        item.Category = catCause;
                        SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 16).Text);

                        SetAttributeWithLogging(log, item, attSortOrder, order);

                        item.Relationship_AddItem(risk, "", Relationship.RelationshipDirection.AtoB);
                    }
                }
                // cause-control
                for (int row = 48; row < 59; row++)
                {
                    string name = XL1.Sheets[sheet].Cells(row, 3).Text;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                        extId = _causeControlsId + $"{order:D2}"; 

                        item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name.Substring(0, Math.Min(50, name.Length)), false);
                        item.ExternalId = extId;
                        item.Description = name;
                        item.Category = catCauseControl;
                        SetAttributeWithLogging(log, item, attOwner, XL1.Sheets[sheet].Cells(row, 9).Text.Trim());
                        SetAttributeWithLogging(log, item, attControlOpinion, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 10).Text.Trim()));
                        SetAttributeWithLogging(log, item, attBasisOfOpinion, XL1.Sheets[sheet].Cells(row, 11).Text.Trim());

                        SetAttributeWithLogging(log, item, attSortOrder, order);
                    }
                }
                // cause-action
                for (int row = 48; row < 59; row++)
                {
                    string name = XL1.Sheets[sheet].Cells(row, 14).Text;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        order = GetInt(XL1.Sheets[sheet].Cells(row, 13).Text.Trim());
                        extId = _causeControlActionsId + $"{order:D2}";

                        item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name.Substring(0, Math.Min(50, name.Length)), false);
                        item.ExternalId = extId;
                        item.Description = name;
                        item.Category = catCauseAction;

                        SetAttributeWithLogging(log, item, attOwner, XL1.Sheets[sheet].Cells(row, 21).Text.Trim());
                        SetAttributeWithLogging(log, item, attPriority, XL1.Sheets[sheet].Cells(row, 22).Text.Trim());
                        SetAttributeWithLogging(log, item, attBaseline, XL1.Sheets[sheet].Cells(row, 23).Text.Trim());
                        SetAttributeWithLogging(log, item, attRevision, XL1.Sheets[sheet].Cells(row, 24).Text.Trim());
                        SetAttributeWithLogging(log, item, attPercComplete, XL1.Sheets[sheet].Cells(row, 25).Text.Trim().Replace("%", ""));
                        SetAttributeWithLogging(log, item, attStatus, XL1.Sheets[sheet].Cells(row, 26).Text.Trim());

                        SetAttributeWithLogging(log, item, attSortOrder, order);
                    }
                }

                // consequences
                for (int row = 23; row < 43; row += 2)
                {
                    string name = XL1.Sheets[sheet].Cells(row, 41).Text;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        order = GetInt(XL1.Sheets[sheet].Cells(row, 40).Text.Trim());
                        extId = _consequenceId + $"{order:D2}";

                        item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name.Substring(0, Math.Min(50, name.Length)), false);
                        item.ExternalId = extId;
                        item.Description = name;
                        item.Category = catConsequence;
                        SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 56).Text);

                        SetAttributeWithLogging(log, item, attSortOrder, order);

                        item.Relationship_AddItem(risk, "", Relationship.RelationshipDirection.BtoA);
                    }
                }
                // consequence-control
                for (int row = 48; row < 59; row++)
                {
                    string name = XL1.Sheets[sheet].Cells(row, 33).Text;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        order = GetInt(XL1.Sheets[sheet].Cells(row, 32).Text.Trim());
                        extId = _consequenceControlsId + $"{order:D2}";

                        item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name.Substring(0, Math.Min(50, name.Length)), false);
                        item.ExternalId = extId;
                        item.Description = name;
                        item.Category = catConsequenceControl;
                        SetAttributeWithLogging(log, item, attOwner, XL1.Sheets[sheet].Cells(row, 39).Text.Trim());
                        SetAttributeWithLogging(log, item, attControlOpinion, LookupControlOpinion (XL1.Sheets[sheet].Cells(row, 40).Text.Trim()));
                        SetAttributeWithLogging(log, item, attBasisOfOpinion, XL1.Sheets[sheet].Cells(row, 41).Text.Trim());

                        SetAttributeWithLogging(log, item, attSortOrder, order);
                    }
                }
                // consequence-action
                for (int row = 48; row < 59; row++)
                {
                    string name = XL1.Sheets[sheet].Cells(row, 45).Text;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        order = GetInt(XL1.Sheets[sheet].Cells(row, 32).Text.Trim());
                        extId = _consequenceControlActionsId + $"{order:D2}";

                        item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name.Substring(0, Math.Min(50, name.Length)), false);
                        item.ExternalId = extId;
                        item.Description = name;
                        item.Category = catConsequenceAction;

                        SetAttributeWithLogging(log, item, attOwner, XL1.Sheets[sheet].Cells(row, 53).Text.Trim());
                        SetAttributeWithLogging(log, item, attPriority, XL1.Sheets[sheet].Cells(row, 54).Text.Trim());
                        SetAttributeWithLogging(log, item, attBaseline, XL1.Sheets[sheet].Cells(row, 55).Text.Trim());
                        SetAttributeWithLogging(log, item, attRevision, XL1.Sheets[sheet].Cells(row, 56).Text.Trim());
                        SetAttributeWithLogging(log, item, attPercComplete, XL1.Sheets[sheet].Cells(row, 57).Text.Trim().Replace("%", ""));
                        SetAttributeWithLogging(log, item, attStatus, XL1.Sheets[sheet].Cells(row, 58).Text.Trim());

                        SetAttributeWithLogging(log, item, attSortOrder, order);
                    }
                }

                for (int row = 9; row <= 17; row += 2)
                {
                    string name = XL1.Sheets[sheet].Cells(row, 22).Text;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        order = counterEWI++;
                        extId = _ewiId + $"{order:D2}";

                        item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name.Substring(0, Math.Min(50, name.Length)), false);
                        item.ExternalId = extId;
                        item.Description = name;
                        item.Category = catEWI;

                        SetAttributeWithLogging(log, item, attLinkedControlsTypes, XL1.Sheets[sheet].Cells(row, 19).Text);
                        SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 21).Text);

                        SetAttributeWithLogging(log, item, attPrior, XL1.Sheets[sheet].Cells(row, 35).Text);
                        SetAttributeWithLogging(log, item, attCurrent, XL1.Sheets[sheet].Cells(row, 37).Text);

                        SetAttributeWithLogging(log, item, attSortOrder, order);

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

            wbBowTie.Close(false);
            XL1.Quit();
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
    }
}
