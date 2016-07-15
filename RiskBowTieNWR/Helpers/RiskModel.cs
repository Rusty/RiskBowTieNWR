using System;
using System.Collections.Generic;
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
        private const string _risk = "Risk";
        private const string _ewi = "Early Warning Indicators";
        private const string _cause = "Causes";
        private const string _causeControls = "Cause Controls";
        private const string _causeControlActions = "Cause Control Actions";
        private const string _consequence = "Consequences";
        private const string _consequenceControls = "Consequence Controls";
        private const string _consequenceActions = "Consequence Control Actions";
        private static readonly string [] _categoryNames = {_causeControlActions, _causeControls, _cause, _risk, _consequence, _consequenceControls, _consequenceActions, _ewi};

        private const string _attrOwner = "Owner.";
        private const string _attrBasisOfOpinion = "Basis of Opinion";
        private const string _attrLinkedControls = "LinkedControls";
        private const string _attrLinkedControlsTypes = "LinkedControlsTypes";
        private static readonly string[] _textFields = { _attrOwner, _attrBasisOfOpinion, _attrLinkedControls, _attrLinkedControlsTypes };

        private const string _attrBaseline = "Base-line";
        private const string _attrRevision = "Revised";
        private static readonly string[] _dateFields = { _attrBaseline, _attrRevision };

        private const string _attrPercComplete = "% Complete";
        private const string _attrOrder = "SortOrder";
        private const string _attrPrior = "Prior";
        private const string _attrCurrent = "Current";
        private static readonly string[] _numberFields = { _attrPercComplete, _attrOrder, _attrPrior, _attrCurrent };

        private const string _attrControlOpinion = "Control Opinion";
        private const string _attrPriority = "Priority";
        private const string _attrStatus = "Status";
        private static readonly string[] _listFields = { _attrControlOpinion, _attrPriority, _attrStatus };

        private static readonly string[] _listCauses = { };
        private static readonly string[] _listCausesControls = { _attrOwner, _attrControlOpinion, _attrBasisOfOpinion };
        private static readonly string[] _listCausesActions = { _attrOwner, _attrPriority, _attrBaseline, _attrBaseline, _attrPercComplete, _attrStatus };
        private static readonly string[] _listConsequenses = { };
        private static readonly string[] _listConsequensesControls = { _attrOwner, _attrControlOpinion, _attrBasisOfOpinion };
        private static readonly string[] _listConsequensesActions = { _attrOwner, _attrPriority, _attrBaseline, _attrBaseline, _attrPercComplete, _attrStatus };
        private static readonly string[] _listEWI = { };


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
                    story.Attribute_Add(a, SC.API.ComInterop.Models.Attribute.AttributeType.Text);
            }
            foreach (var a in _dateFields)
            {
                if (story.Attribute_FindByName(a) == null)
                    story.Attribute_Add(a, SC.API.ComInterop.Models.Attribute.AttributeType.Date);
            }
            foreach (var a in _numberFields)
            {
                if (story.Attribute_FindByName(a) == null)
                    story.Attribute_Add(a, SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            }
            foreach (var a in _listFields)
            {
                if (story.Attribute_FindByName(a) == null)
                    story.Attribute_Add(a, SC.API.ComInterop.Models.Attribute.AttributeType.List);
            }

        }

        public static async void ProcessBowTies(SharpCloudApi sc, string teamId, string portfolioId, string controlId, string temaplateId, Logger log)
        {
            var portfolioStory = sc.LoadStory(portfolioId);
            var controlsStory = sc.LoadStory(controlId);

            foreach (var teamStory in sc.StoriesTeam(teamId))
            {
                if (teamStory.Id != portfolioId && teamStory.Id != temaplateId && teamStory.Id != controlId)
                {
                    log.Log($"Reading from '{teamStory.Name}'");
                    await Task.Delay(100);

                    var riskItem = portfolioStory.Item_FindByName(teamStory.Name);
                    if (riskItem == null)
                        riskItem = portfolioStory.Item_AddNew(teamStory.Name, false);


                    try
                    {
                        var story = sc.LoadStory(teamStory.Id);
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

                        GetAttributeData(riskItem, portfolioStory, story, new[] { "Likelihood", "Impact" });

                        foreach (var item1 in story.Items)
                        {
                            if (item1.Category.Name == "Controls")
                            {
                                var cItem = controlsStory.Item_FindByName(item1.Name);
                                if (cItem == null)
                                    cItem = controlsStory.Item_AddNew(item1.Name);

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
                        log.LogError(string.Format($"Could not open {riskItem.ExternalId}"));
                    }
                }
            }

            portfolioStory.Save();
            controlsStory.Save();

        }

        private static void GetAttributeData(Item item, Story storyP, Story story, string[] strings)
        {
            foreach (var item2 in story.Items)
            {
                if (item2.Category.Name == "Risk")
                {
                    foreach (var s in strings)
                    {
                        item.SetAttributeValue(storyP.Attribute_FindByName(s),
                            item2.GetAttributeValueAsText(story.Attribute_FindByName(s)));
                    }
                }
            }
        }

        private static void LoadPanelData(Item item, Story story, string category, string[] attributes)
        {
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
                table1.SetValue(0, col++, attribute);
            }

            int row = 1;
            foreach (var item2 in story.Items)
            {
                if (item2.Category.Name == category)
                {
                    col = 0;
                    table1.SetValue(row, col++, item2.Name);
                    foreach (var attribute in attributes)
                        table1.SetValue(row, col++, item2.GetAttributeValueAsText(story.Attribute_FindByName(attribute)));
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

            // get atttributes
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


            var XL1 = new Application();
            var pathMlstn = XLFilename;
            log.Log($"Opening Excel Doc " + pathMlstn);
            var wbBowTie = XL1.Workbooks.Open(pathMlstn);
            var sheet = 1;

            Item risk = story.Item_FindByExternalId(_riskId) ?? story.Item_AddNew(XL1.Sheets[sheet].Cells(3, 4).Text, false);
            risk.ExternalId = _riskId;
            risk.Description = XL1.Sheets[sheet].Cells(3, 19).Text;
            risk.Category = catRisk;
            risk.SetAttributeValue(story.Attribute_FindByName("Impact"), "3");
            risk.SetAttributeValue(story.Attribute_FindByName("Likelihood"), "3 - Medium");

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
                        SetAttributeWithLogging(log, item, attControlOpinion, XL1.Sheets[sheet].Cells(row, 10).Text.Trim());
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
                        SetAttributeWithLogging(log, item, attControlOpinion, XL1.Sheets[sheet].Cells(row, 40).Text.Trim());
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
                item.SetAttributeValue(att, value);
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
