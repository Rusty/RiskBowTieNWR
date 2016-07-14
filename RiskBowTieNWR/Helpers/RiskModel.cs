using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using SC.API.ComInterop.Models;

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
        private const string _consequenceControls = "Consequence Control";
        private const string _consequenceControlActions = "Consequence Control Actions";
        private static readonly string [] _categoryNames = {_causeControlActions, _causeControls, _cause, _risk, _consequence, _consequenceControls, _consequenceControlActions, _ewi};

        private const string _attrOwner = "Owner.";
        private const string _attrBasisOfOpinion = "Basis of Opinion";
        private const string _attrLinkedControls = "LinkedControls";
        private static readonly string[] _textFields = { _attrOwner, _attrBasisOfOpinion, _attrLinkedControls };

        private const string _attrBaseline = "Base-line";
        private const string _attrRevision = "Revised";
        private static readonly string[] _dateFields = { _attrBaseline, _attrRevision };

        private const string _attrPercComplete = "% Complete";
        private const string _attrOrder = "SortOrder";
        private static readonly string[] _numberFields = { _attrPercComplete, _attrOrder };

        private const string _attrControlOpinion = "Control Opinion";
        private const string _attrPriority = "Priority";
        private const string _attrStatus = "Status";
        private static readonly string[] _listFields = { _attrControlOpinion, _attrPriority, _attrStatus };

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
            var catConsequenceAction = story.Category_FindByName(_consequenceControlActions);

            // get atttributes
            var attOwner = story.Attribute_FindByName(_attrOwner);
            var attBasisOfOpinion = story.Attribute_FindByName(_attrBasisOfOpinion);
            var attLinkedControls = story.Attribute_FindByName(_attrLinkedControls);
            var attBaseline = story.Attribute_FindByName(_attrBaseline);
            var attRevision = story.Attribute_FindByName(_attrRevision);
            var attPercComplete = story.Attribute_FindByName(_attrPercComplete);
            var attSortOrder = story.Attribute_FindByName(_attrOrder);
            var attControlOpinion = story.Attribute_FindByName(_attrControlOpinion);
            var attPriority = story.Attribute_FindByName(_attrPriority);
            var attStatus = story.Attribute_FindByName(_attrStatus);


            var XL1 = new Application();
            var pathMlstn = XLFilename;
            log.Log($"Opening Excel Doc " + pathMlstn);
            var wbBowTie = XL1.Workbooks.Open(pathMlstn);
            var sheet = 1;

            Item risk = story.Item_AddNew(XL1.Sheets[sheet].Cells(3, 4).Text, false);
            risk.Description = XL1.Sheets[sheet].Cells(3, 19).Text;
            risk.Category = catRisk;
            risk.SetAttributeValue(story.Attribute_FindByName("Impact"), "3");
            risk.SetAttributeValue(story.Attribute_FindByName("Likelihood"), "3 - Medium");

            Item item;
            string extId;
            int order;
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


                for (int row = 9; row < 17; row++)
                {
                    string name = XL1.Sheets[sheet].Cells(row, 22).Text;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        order = GetInt(XL1.Sheets[sheet].Cells(row, 32).Text.Trim());
                        extId = _ewiId + $"{order:D2}";

                        item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name.Substring(0, Math.Min(50, name.Length)), false);
                        item.ExternalId = extId;
                        item.Description = name;
                        item.Category = catEWI;

                        SetAttributeWithLogging(log, item, attSortOrder, order);
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

            }

            wbBowTie.Close(false);
            XL1.Quit();
        }

        private static void SetAttributeWithLogging(Logger log, Item item, SC.API.ComInterop.Models.Attribute att, int value)
        {
            SetAttributeWithLogging(log, item, att, $"{value:D2}");
        }

        private static void SetAttributeWithLogging(Logger log, Item item, SC.API.ComInterop.Models.Attribute att, string value)
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
