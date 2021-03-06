﻿using System;
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
        private static readonly string[] _categoryNames = { _causeControlActions, _causeControls, _cause, _risk, _consequence, _consequenceControls, _consequenceActions, _ewi };

        private const string _attrVersion = "Version";
        private const string _attrRiskOwner = "Risk Owner";
        private const string _attrControlOwner = "Control Owner";
        private const string _attrActionOwner = "Action Owner";
        private const string _attrManager = "Risk Manager";
        private const string _attrLinkedControls = "Linked Controls";

        private const string _attrControlIndustryPartner = "Z. Industry Partner";
        private const string _attrControlOpinionRationale = "Z. Opinion Rationale";
        private const string _attrControlOpinionSource = "Z. Opinion Source";

        private const string _attrLinkedControlsTypes = "Linked Control Type";
        private const string _attrRationale = "Rationale (Overall)";
        private const string _attrRationaleSafety = "Rationale (SHE)";
        private const string _attrRationalePerformance = "Rationale (Performance)";
        private const string _attrRationaleValue = "Rationale (Finance/Value)";
        private const string _attrRationalePolitical = "Rationale (Political/Reputation)";
        private const string _attEWISparkVals = "Z. EWISparkVals";
        private const string _attEWISparkValsTolerance = "Z. EWISparkValsTolerance";
        private const string _attEWISparkValsActual = "Z. EWISparkValsActual";

        private static readonly string[] _textFields = { _attrVersion, _attrRiskOwner, _attrControlOwner, _attrActionOwner, _attrManager, _attrLinkedControls, _attrLinkedControlsTypes, _attrRationale,
            _attrRationaleSafety, _attrRationalePerformance, _attrRationaleValue, _attrRationalePolitical,_attrControlIndustryPartner,_attrControlOpinionRationale,
            _attrControlOpinionSource,_attEWISparkVals,_attEWISparkValsTolerance,_attEWISparkValsActual, _attrRiskGrossToNet,_attrRiskNetToTarget};

        private const string _attrBaseline = "Action Baseline Due Date";
        private const string _attrRevision = "Action Revised Due Date";
        private const string _attrlastUpdate = "Last Update";
        private static readonly string[] _dateFields = { _attrBaseline, _attrRevision, _attrRiskExposureFrame };

        private const string _attrPercComplete = "Action % Complete";
        private const string _attrOrder = "SortOrder";
        //        private const string _attrPrior = "Prior";
        private const string _attrCurrent = "Current";
        private const string _attrTolerance = "Tolerance";
        private static readonly string[] _numberFields = { _attrPercComplete, _attrOrder, _attrCurrent, _attrTolerance };

        // control lib only
        private const string _attrRelatedRisks = "Risk Count";


        private const string _attrControlOpinion = "Control Opinion";
        private const string _attrBasisOfOpinion = "Basis of Opinion";
        private const string _attrPriority = "Action Priority";
        private const string _attrStatus = "Action Status";
        private const string _attrClassification = "Classification";
        private const string _attrImpactedArea = "Key Scorecard Area Impacted";
        private const string _attrControlRating = "Overall Control Rating";
        private const string _attrRiskLevel = "Risk Level";
        private const string _attrRiskAppetiteSafety = "Above SHE Risk Appetite";
        private const string _attrRiskAppetitePerformance = "Above Performance Risk Appetite";
        private const string _attrRiskAppetiteValue = "Above Finance/Value Risk Appetite";
        private const string _attrRiskAppetitePolitical = "Above Political/Reputation Risk Appetite";
        private const string _attrReportingPriority = "Reporting Priority";
        private const string _attrDirectorate = "Directorate";
        private const string _attrSubDirectorate = "Z .SubDirectorate";
        private const string _attrGrossRating = "Gross Rating";
        private const string _attrTargetRating = "Target Rating";
        private const string _attrWithinTolerance = "Within Tolerance";
        private const string _attrRiskCategory = "Z. Risk Category";
        private const string _attrRiskType = "Z. Risk Type";
        private const string _attrControlIndustry = "Z. Industry";
        private const string _attrControlIndustryVisibility = "Z. Industry Visibility";

        private const string _attrRiskM1Impact = "Z. M1 Impact";
        private const string _attrRiskM1Likelihood = "Z. M1 Likelihood";
        private const string _attrRiskM1Finance = "Z. M1 Finance";
        private const string _attrRiskM1Rating = "Z. M1 Rating";

        private const string _attrRiskM2Impact = "Z. M2 Impact";
        private const string _attrRiskM2Likelihood = "Z. M2 Likelihood";
        private const string _attrRiskM2Finance = "Z. M2 Finance";
        private const string _attrRiskM2Rating = "Z. M2 Rating";

        private const string _attrRiskM3Impact = "Z. M3 Impact";
        private const string _attrRiskM3Likelihood = "Z. M3 Likelihood";
        private const string _attrRiskM3Finance = "Z. M3 Finance";
        private const string _attrRiskM3Rating = "Z. M3 Rating";

        private const string _attrRiskM4Impact = "Z. M4 Impact";
        private const string _attrRiskM4Likelihood = "Z. M4 Likelihood";
        private const string _attrRiskM4Finance = "Z. M4 Finance";
        private const string _attrRiskM4Rating = "Z. M4 Rating";

        private const string _attrRiskM5Impact = "Z. M5 Impact";
        private const string _attrRiskM5Likelihood = "Z. M5 Likelihood";
        private const string _attrRiskM5Finance = "Z. M5 Finance";
        private const string _attrRiskM5Rating = "Z. M5 Rating";

        private const string _attrRiskGrossToNet = "Z. Gross to Net";
        private const string _attrRiskNetToTarget = "Z. Net to Target";


        private static readonly string[] _listFields = { _attrControlOpinion, _attrBasisOfOpinion, _attrPriority, _attrStatus, _attrClassification, _attrImpactedArea, _attrControlRating, _attrRiskLevel,
            _attrRiskAppetiteSafety, _attrRiskAppetitePerformance, _attrRiskAppetiteValue,_attrRiskAppetitePolitical, _attrReportingPriority, _attrDirectorate,_attrRiskCategory,
            _attrGrossRating,  _attrTargetRating, _attrWithinTolerance, _attrRiskType,_attrControlIndustry,_attrControlIndustryVisibility,_attrRiskM1Impact, _attrRiskM1Likelihood,_attrRiskM1Finance,_attrRiskM1Rating,
            _attrRiskM2Impact, _attrRiskM2Likelihood,_attrRiskM2Finance,_attrRiskM2Rating,
        _attrRiskM3Impact, _attrRiskM3Likelihood,_attrRiskM3Finance,_attrRiskM3Rating,
        _attrRiskM4Impact, _attrRiskM4Likelihood,_attrRiskM4Finance,_attrRiskM4Rating,
        _attrRiskM5Impact, _attrRiskM5Likelihood,_attrRiskM5Finance,_attrRiskM5Rating,
        };

        private const string _attrLikelihood = "Likelihood (Overall)";
        private const string _attrImpact = "Impact (Overall)";
        private const string _attrLikelihoodSafety = "Likelihood (SHE)";
        private const string _attrImpactSafety = "Impact (SHE)";
        private const string _attrLikelihoodPerformance = "Likelihood (Performance)";
        private const string _attrImpactPerformance = "Impact (Performace)";
        private const string _attrLikelihoodValue = "Likelihood (Finance/Value)";
        private const string _attrImpactValue = "Impact (Finance/Value)";
        private const string _attrLikelihoodPolitical = "Likelihood (Political/Reputation)";
        private const string _attrImpactPolitical = "Impact (Political/Reputation)";
        private const string _attrGrossImpact = "Gross Impact";
        private const string _attrGrossLkelihood = "Gross Likelihood";
        private const string _attrGrossFinance = "Gross Finance";
        private const string _attrTargetImpact = "Target Impact";
        private const string _attrTargetLkelihood = "Target Likelihood";
        private const string _attrTargetFinance = "Target Finance";
        private const string _attrRiskExposureFrame = "Z. Exposure Frame";
       
        private static readonly string[] _listRiskFields = { _attrLikelihood, _attrImpact, _attrLikelihoodSafety, _attrImpactSafety, _attrLikelihoodPerformance, _attrImpactPerformance,
            _attrLikelihoodValue, _attrImpactValue, _attrLikelihoodPolitical, _attrImpactPolitical,
             _attrGrossImpact, _attrGrossLkelihood, _attrGrossFinance, _attrTargetImpact, _attrTargetLkelihood, _attrTargetFinance};
        



        private static readonly string _sort = "#";
        private static readonly string _name = "Name";
        private static readonly string[] _listCauses = { };
        private static readonly string[] _listCausesControls = { _attrControlOwner, _attrControlOpinion, _attrBasisOfOpinion };
        private static readonly string[] _listCausesActions = { _attrActionOwner, _attrPriority, _attrBaseline, _attrPercComplete, _attrStatus };
        private static readonly string[] _listConsequenses = { };
        private static readonly string[] _listConsequensesControls = { _attrControlOwner, _attrControlOpinion, _attrBasisOfOpinion };
        private static readonly string[] _listConsequensesActions = { _attrActionOwner, _attrPriority, _attrBaseline, _attrPercComplete, _attrStatus };
        private static readonly string[] _listEWI = { };

        private static readonly double[] _widthCauses = { 10, 90 };
        private static readonly double[] _widthCausesControls = { 10, 30, 30, 20, 10 };
        private static readonly double[] _widthCausesActions = { 10, 40, 10, 10, 10, 10, 10 };
        private static readonly double[] _widthConsequenses = { 10, 90 };
        private static readonly double[] _widthConsequensesControls = { 10, 30, 30, 20, 10 };
        private static readonly double[] _widthConsequensesActions = { 10, 40, 10, 10, 10, 10, 10 };
        private static readonly double[] _widthEWI = { 10, 90 };


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


        // relationship fields
        private const string _attrControlOpinionRels = "Control Opinion";
        private const string _attrControlBasisOfOpinionRels = "Basis of Opinion";
        private static readonly string[] _listFieldsRels = { _attrControlBasisOfOpinionRels, _attrControlOpinionRels };

        private const string _attrControlOwnerRels = "Control Owner";
        private const string _attrOpinionRationalerRels = "Opinion Rationale";
        private static readonly string[] _textFieldsRels = { _attrControlOwnerRels, _attrOpinionRationalerRels };

        private static Dictionary<string, Item> _sharedControlDictionary = null;


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

            // add the relationship attributes
            foreach (var a in _textFieldsRels)
            {
                if (story.RelationshipAttribute_FindByName(a) == null)
                {
                    log.Log($"Adding Relationship Text Attribute '{a}'");
                    story.RelationshipAttribute_Add(a, RelationshipAttribute.RelationshipAttributeType.Text);
                }
            }
            foreach (var a in _listFieldsRels)
            {
                if (story.RelationshipAttribute_FindByName(a) == null)
                {
                    log.Log($"Adding Relationship List Attribute '{a}'");
                    story.RelationshipAttribute_Add(a, RelationshipAttribute.RelationshipAttributeType.List);
                }
            }
        }

        public static void EnsureControlStoryHasRightStructure(Story controlstory, Logger log)
        {
            foreach (var a in _listFieldsRels)
            {
                if (controlstory.RelationshipAttribute_FindByName(a) == null)
                {
                    log.Log($"Adding Relationship List Attribute '{a}'");
                    controlstory.RelationshipAttribute_Add(a, RelationshipAttribute.RelationshipAttributeType.List);
                }
            }
            if (controlstory.Attribute_FindByName(_attrRelatedRisks) == null)
                controlstory.Attribute_Add(_attrRelatedRisks, Attribute.AttributeType.Numeric);

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

        public static string LookupYesNoRiskLabel(string l)
        {
            switch (l.ToLower().Trim())
            {
                case "no":
                    return "No";
            }
            return "Yes"; // everything else
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

            var words = str.Split('_');

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
            var attSubCat = portfolioStory.Attribute_FindByName("Risk Sub Category") ??
                               portfolioStory.Attribute_Add("Risk Sub Category", Attribute.AttributeType.List);


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
            var attDirectorateControlStory = controlsStory.Attribute_FindByName(_attrDirectorate) ??
                               controlsStory.Attribute_Add(_attrDirectorate, Attribute.AttributeType.List);
            var attControlOwnerControlStory = controlsStory.Attribute_FindByName(_attrControlOwner) ??
                               controlsStory.Attribute_Add(_attrControlOwner, Attribute.AttributeType.List);


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
                itm.RemoveAttributeValue(attDirectorateControlStory);
                itm.RemoveAttributeValue(attControlOwnerControlStory);

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

                            string catName, subCatName;
                            var split = riskCategoryName.Split('.');
                            if (split.Count() == 1)
                            {
                                catName = split[0];
                                subCatName = "";
                            }
                            else // maybe multiple sub cats
                            {
                                catName = split[0];
                                subCatName = riskCategoryName.Substring(catName.Length+1);
                            }

                            var riskCategory = portfolioStory.Category_FindByName(catName) ??
                                               portfolioStory.Category_AddNew(catName);
                            riskItem.Category = riskCategory;

                            riskItem.SetAttributeValue(attSubCat, subCatName); // set the sub category attribute
                            // TODO Set the real sub category when the SDK supports this.

                            // copy and assigned risk tags
                            foreach (var t in riskItemSource.Tags)
                                riskItem.Tag_AddNew(t.Text);


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
                                                       story.Attribute_Add(_attrControlOpinion, Attribute.AttributeType.List);

                            var attRiskLevelControlsRiskStory = story.Attribute_FindByName(_attrRiskLevel) ??
                                                       story.Attribute_Add(_attrRiskLevel, Attribute.AttributeType.List);

                            var attOverallControlRatingControlsRiskStory = story.Attribute_FindByName(_attrControlRating) ??
                                                       story.Attribute_Add(_attrControlRating, Attribute.AttributeType.List);

                            var attBasisOfOpinionRiskStory = story.Attribute_FindByName(_attrBasisOfOpinion) ??
                                                       story.Attribute_Add(_attrBasisOfOpinion, Attribute.AttributeType.List);

                            var attScorecardAreaRiskStory = story.Attribute_FindByName(_attrImpactedArea) ??
                                                       story.Attribute_Add(_attrImpactedArea, Attribute.AttributeType.List);

                            var attControlOwnerRiskStory = story.Attribute_FindByName(_attrControlOwner) ??
                                                       story.Attribute_Add(_attrControlOwner, Attribute.AttributeType.List);


                            foreach (var itemControlSource in story.Items.Where( i =>
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
                                // add risk directorate
                                AddAttributeButCheckForDiffernce(riskItemSource, attDirectorate, itemControlDestination, attDirectorateControlStory);
                                // add risk directorate
                                AddAttributeButCheckForDiffernce(itemControlSource, attControlOwnerRiskStory, itemControlDestination, attControlOwnerControlStory);

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


        private static void LoadPanelData(Item item, Story story, string category, string[] attributes, double[] widths)
        {
            if (attributes.Length+2 != widths.Length)
            {
                throw new Exception("attributes and widths must match");
            }

            var sortby = story.Attribute_FindByName("SortOrder");
            
            // add attributes so it won't blow up below
            foreach (var attribute in attributes)
            {
                if (story.Attribute_FindByName(attribute) == null)
                    story.Attribute_Add(attribute, Attribute.AttributeType.List);
            }

            var table1 = new HTMLTable(attributes.Length + 2);
            int col = 0;
            table1.SetValue(0, col++, "#");
            table1.SetValue(0, col++, "Name");
            foreach (var attribute in attributes)
            {
                table1.SetValue(0, col++, attribute);
            }

            int row = 1;
            foreach (var item2 in story.Items.OrderBy(i => i.GetAttributeValueAsDouble(sortby)))
            {
                if (item2.Category.Name == category)
                {
                    col = 0;
                    table1.SetValue(row, col++, item2.GetAttributeValueAsDouble(sortby).ToString());
                    table1.SetValue(row, col++, item2.Name);
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

            // set the col width
            for (int c=0;c<widths.Length; c++)
            {
                table1.SetColWidth(c, widths[c]);
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

            var id = XL1.Sheets[1].Cells(14, 4).Text;

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

            XL1.Sheets[1].Cells[14, 4] = Id;

            wbBowTie.Close(true);
            XL1.Quit();
        }

        private static void RemoveControlRels(Story story, Story controlStory, Logger log)
        {
            try
            {
                // create collection of unique element Id's
                var listIds = new List<string>();
                foreach (var i in story.Items)
                {
                    listIds.Add(i.Id);
                }

                var toDelete = new List<Relationship>();
                foreach (var r in controlStory.Relationships)
                {
                    var rel = r.AsRelationship;
                    if (listIds.Contains(rel.Element1ID.ToString()) || listIds.Contains(rel.Element2ID.ToString()))
                        toDelete.Add(r);
                }
                foreach (var rel in toDelete)
                    controlStory.Relationship_DeleteById(rel.Id);
            }
            catch (Exception ex)
            {
                log.LogError(ex);
            }
        }
        
        public static void CreateStoryFromXLTemplate(Story story, Story controlStory,string XLFilename, Logger log, bool deleteItems, bool deleteRels, bool verbose,string excelVersion)
        {
            if (verbose) log.Log("Removing old cross story relationships to Controls Library");
            RemoveControlRels(story, controlStory, log);

            _sharedControlDictionary = new Dictionary<string, Item>();

            if (verbose) log.Log("Checking structure");
            EnsureStoryHasRightStructure(story, log);
            EnsureControlStoryHasRightStructure(controlStory, log);

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
            var attRiskOwner = story.Attribute_FindByName(_attrRiskOwner);
            var attControlOwner = story.Attribute_FindByName(_attrControlOwner);
            var attActionOwner = story.Attribute_FindByName(_attrActionOwner);
            var attBasisOfOpinion = story.Attribute_FindByName(_attrBasisOfOpinion);
            var attLinkedControls = story.Attribute_FindByName(_attrLinkedControls);
            var attControlIndustry = story.Attribute_FindByName(_attrControlIndustry);
            var attControlIndustryPartner = story.Attribute_FindByName(_attrControlIndustryPartner);
            var attControlIndustryVisibility = story.Attribute_FindByName(_attrControlIndustryVisibility);
            var attControlOpinionRationale = story.Attribute_FindByName(_attrControlOpinionRationale);
            var attControlOpinionSource = story.Attribute_FindByName(_attrControlOpinionSource);


            var attLinkedControlsTypes = story.Attribute_FindByName(_attrLinkedControlsTypes);
            var attBaseline = story.Attribute_FindByName(_attrBaseline);
            var attRevision = story.Attribute_FindByName(_attrRevision);
            var attPercComplete = story.Attribute_FindByName(_attrPercComplete);
            var attSortOrder = story.Attribute_FindByName(_attrOrder);
            //var attPrior = story.Attribute_FindByName(_attrPrior);
            var attCurrent = story.Attribute_FindByName(_attrCurrent);

            var attTolerance = story.Attribute_FindByName(_attrTolerance);
            var attWithinTolerance = story.Attribute_FindByName(_attrWithinTolerance);
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
            var attRiskType = story.Attribute_FindByName(_attrRiskType);
            var attrRiskExposureFrame = story.Attribute_FindByName(_attrRiskExposureFrame);
            var attrRiskCategory = story.Attribute_FindByName(_attrRiskCategory);
            var attEWISparkVals = story.Attribute_FindByName(_attEWISparkVals);
            var attEWISparkValsActual = story.Attribute_FindByName(_attEWISparkValsActual);
            var attEWISparkValsTolerance = story.Attribute_FindByName(_attEWISparkValsTolerance);
            var attSubDirectorate = story.Attribute_FindByName(_attrSubDirectorate);

            var attRiskM1Impact = story.Attribute_FindByName(_attrRiskM1Impact);
            var attRiskM1Likelihood = story.Attribute_FindByName(_attrRiskM1Likelihood);
            var attRiskM1Finance = story.Attribute_FindByName(_attrRiskM1Finance);
            var attRiskM1Rating = story.Attribute_FindByName(_attrRiskM1Rating);

            var attRiskM2Impact = story.Attribute_FindByName(_attrRiskM2Impact);
            var attRiskM2Likelihood = story.Attribute_FindByName(_attrRiskM2Likelihood);
            var attRiskM2Finance = story.Attribute_FindByName(_attrRiskM2Finance);
            var attRiskM2Rating = story.Attribute_FindByName(_attrRiskM2Rating);

            var attRiskM3Impact = story.Attribute_FindByName(_attrRiskM3Impact);
            var attRiskM3Likelihood = story.Attribute_FindByName(_attrRiskM3Likelihood);
            var attRiskM3Finance = story.Attribute_FindByName(_attrRiskM3Finance);
            var attRiskM3Rating = story.Attribute_FindByName(_attrRiskM3Rating);

            var attRiskM4Impact = story.Attribute_FindByName(_attrRiskM4Impact);
            var attRiskM4Likelihood = story.Attribute_FindByName(_attrRiskM4Likelihood);
            var attRiskM4Finance = story.Attribute_FindByName(_attrRiskM4Finance);
            var attRiskM4Rating = story.Attribute_FindByName(_attrRiskM4Rating);

            var attRiskM5Impact = story.Attribute_FindByName(_attrRiskM5Impact);
            var attRiskM5Likelihood = story.Attribute_FindByName(_attrRiskM5Likelihood);
            var attRiskM5Finance = story.Attribute_FindByName(_attrRiskM5Finance);
            var attRiskM5Rating = story.Attribute_FindByName(_attrRiskM5Rating);

            var attRiskGrossToNet = story.Attribute_FindByName(_attrRiskGrossToNet);
            var attRiskNetToTarget = story.Attribute_FindByName(_attrRiskNetToTarget);

      


        // relationship attributes
            var attBasisOfOpinionRels = story.RelationshipAttribute_FindByName(_attrControlBasisOfOpinionRels);
            var attControlOwnerRels = story.RelationshipAttribute_FindByName(_attrControlOwnerRels);
            var attControlOpinionRels = story.RelationshipAttribute_FindByName(_attrControlOpinionRels);
            var attOpinionRationalerRels = story.RelationshipAttribute_FindByName(_attrOpinionRationalerRels);

            var attCBasisOfOpinionRels = controlStory.RelationshipAttribute_FindByName(_attrControlBasisOfOpinionRels);
            var attCControlOwnerRels = controlStory.RelationshipAttribute_FindByName(_attrControlOwnerRels);
            var attCControlOpinionRels = controlStory.RelationshipAttribute_FindByName(_attrControlOpinionRels);

            log.Log($"Using  Version  " + excelVersion);
            if (excelVersion == "4")
            {

                try
                {
                    var sheetNames = new string[] { "ERR", "ERR Cont Sheet 1", "ERR Cont Sheet 2" };

                    var XL1 = new Application();
                    var pathMlstn = XLFilename;
                    log.Log($"Opening Excel Doc " + pathMlstn);
                    var wbBowTie = XL1.Workbooks.Open(pathMlstn);
                    var sheet = sheetNames[0];

                    // validate template is correct version
                    var version = XL1.Sheets["Version Control"].Cells[1, 26].Text;
                    if (version != "SCApproved")
                    {
                        log.Log($"Spreadsheet is not in the approved version, missing 'SCApproved' at Z1 in 'Version Control' ");
                        KillProcessByMainWindowHwnd(XL1.Application.Hwnd);
                        return;
                    }
                    string version2 = XL1.Sheets["ERR"].Cells[1, 14].Text;
                    if (!version2.Contains("v4"))
                    {
                        log.Log($"Spreadsheet is not in the approved version, missing 'v4x' at 'N:1' in 'ERR' ");
                        KillProcessByMainWindowHwnd(XL1.Application.Hwnd);
                        return;
                    }

                    var list = new Dictionary<string, Relationship>();
                    foreach (var rel in story.Relationships)
                        list.Add(rel.Id, rel);

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

                    SetAttributeWithLogging(log, risk, attRiskOwner, XL1.Sheets[sheet].Cells(6, 4).Text);
                    SetAttributeWithLogging(log, risk, attManager, XL1.Sheets[sheet].Cells(7, 4).Text);
                    SetAttributeWithLogging(log, risk, attImapactedArea, LookupRiskLabel(XL1.Sheets[sheet].Cells(8, 4).Text));
                    SetAttributeWithLogging(log, risk, attControlRating, XL1.Sheets[sheet].Cells(39, 30).Text); // new position on template
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
                    SetAttributeWithLogging(log, risk, attAppetiteSafety, LookupYesNoRiskLabel(XL1.Sheets[sheet].Cells(22, 35).Text));
                    SetAttributeWithLogging(log, risk, attRationaleSafety, XL1.Sheets[sheet].Cells(20, 19).Text);

                    SetAttributeWithLogging(log, risk, attLikelihoodPerformance, LookupRiskLabel(XL1.Sheets[sheet].Cells(25, 37).Text));
                    SetAttributeWithLogging(log, risk, attImpactPerformance, LookupRiskLabel(XL1.Sheets[sheet].Cells(25, 35).Text));
                    SetAttributeWithLogging(log, risk, attAppetitePerformace, LookupYesNoRiskLabel(XL1.Sheets[sheet].Cells(27, 35).Text));
                    SetAttributeWithLogging(log, risk, attRationalePerformance, XL1.Sheets[sheet].Cells(25, 19).Text);

                    SetAttributeWithLogging(log, risk, attLikelihoodValue, LookupRiskLabel(XL1.Sheets[sheet].Cells(30, 37).Text));
                    SetAttributeWithLogging(log, risk, attImpactValue, LookupRiskLabel(XL1.Sheets[sheet].Cells(30, 35).Text));
                    SetAttributeWithLogging(log, risk, attAppetiteValue, LookupYesNoRiskLabel(XL1.Sheets[sheet].Cells(32, 35).Text));
                    SetAttributeWithLogging(log, risk, attRationaleValue, XL1.Sheets[sheet].Cells(30, 19).Text);

                    SetAttributeWithLogging(log, risk, attLikelihoodPolitical, LookupRiskLabel(XL1.Sheets[sheet].Cells(35, 37).Text));
                    SetAttributeWithLogging(log, risk, attImpactPolitical, LookupRiskLabel(XL1.Sheets[sheet].Cells(35, 35).Text));
                    SetAttributeWithLogging(log, risk, attAppetitePolitical, LookupYesNoRiskLabel(XL1.Sheets[sheet].Cells(37, 35).Text));
                    SetAttributeWithLogging(log, risk, attRationalePolitical, XL1.Sheets[sheet].Cells(35, 19).Text);

                    SetAttributeWithLogging(log, risk, attLikelihood, LookupRiskLabel(XL1.Sheets[sheet].Cells(2, 78).Text)); //TODO
                    SetAttributeWithLogging(log, risk, attImpact, LookupRiskLabel(XL1.Sheets[sheet].Cells(1, 78).Text)); //TODO
                    SetAttributeWithLogging(log, risk, attRationale, XL1.Sheets[sheet].Cells(40, 19).Text);

                    SetAttributeWithLogging(log, risk, attReportingPriority, GetReportingPriority(0));

                    string tagText = XL1.Sheets["Version Control"].Cells[4, 8].Text;
                    var tags = tagText.Split(',');
                    foreach (var t in tags)
                    {
                        risk.Tag_AddNew(t.Trim());
                    }
                 

                    Item item;
                    string extId;
                    int order;
                    int counterEWI = 1;
                    string name;
                    string desc;
                    string text;
                    int sht;

                    //special code to delete old consequence actions
                    if (deleteItems)
                    {
                        for (int i = 1; i <= 36; i++)
                        {
                            extId = _consequenceControlActionsId + $"{i:D2}";
                            DeleteItemWithLogging(log, story, extId);
                        }

                        DeleteItemWithLogging(log, story, _causeControls + $"{0:D2}");
                        DeleteItemWithLogging(log, story, _consequenceControls + $"{0:D2}");
                    }


                    // data can be in the same place on 3 sheets (continuation sheets)
                    for (sht = 0; sht < 3; sht++)
                    {
                        sheet = sheetNames[sht];

                        try
                        {
                            var testsheet = XL1.Sheets[sheet];
                            if (testsheet == null)
                            {
                                log.Log($"Sheet '{sheet}' does not exist, skipping.");
                                continue;
                            }
                        }
                        catch (Exception exsheet)
                        {
                            log.Log($"Sheet '{sheet}' does not exist, skipping.");
                            continue;
                        }


                        if (verbose) log.Log($"Processing Sheet{sheet}");

                        // cause
                        for (int row = 23; row < 43; row += 2)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 3).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {

                                GetItemNameAndDescription(text, out name, out desc);
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                                extId = _causeId + $"{order:D2}";
                                if (verbose) log.Log($"Processing '{extId}'");

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catCause;
                                SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 16).Text);
                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));
                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                                extId = _causeId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting {extId}");
                            }
                        }
                        // cause-control
                        for (int row = 48; row <= 59; row++)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 3).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                GetItemNameAndDescription(text, out name, out desc);
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                                extId = _causeControlsId + $"{order:D2}";
                                if (verbose) log.Log($"Processing '{extId}'");

                                var words = name.Split(' ');
                                Item itemShared = null;
                                if (words.Any())
                                {
                                    // a shared item exists in the control library
                                    itemShared = controlStory.Item_FindByExternalId(words[0]);
                                }

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catCauseControl;
                                // awlays try to set these
                                SetAttributeWithLogging(log, item, attControlOwner, XL1.Sheets[sheet].Cells(row, 9).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlOpinion, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 10).Text.Trim()));
                                SetAttributeWithLogging(log, item, attBasisOfOpinion, XL1.Sheets[sheet].Cells(row, 11).Text.Trim());

                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));

                                var rel = item.Relationship_AddItem(risk, "", Relationship.RelationshipDirection.AtoB);
                                SetRelAttributeWithLogging(log, rel, attControlOwnerRels, XL1.Sheets[sheet].Cells(row, 9).Text.Trim());
                                SetRelAttributeWithLogging(log, rel, attControlOpinionRels, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 10).Text.Trim()));
                                SetRelAttributeWithLogging(log, rel, attBasisOfOpinionRels, XL1.Sheets[sheet].Cells(row, 11).Text.Trim());

                                if (itemShared != null) // should be a shared control 
                                {
                                    log.Log($"DETECTED SHARED CONTROL '{words[0]}'");
                                    log.Log($"'{itemShared.Name}'");

                                    var rel2 = itemShared.Relationship_AddItem(risk, "RISK", Relationship.RelationshipDirection.AtoB);
                                    SetRelAttributeWithLogging(log, rel2, attCControlOwnerRels, XL1.Sheets[sheet].Cells(row, 9).Text.Trim());
                                    SetRelAttributeWithLogging(log, rel2, attCControlOpinionRels, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 10).Text.Trim()));
                                    SetRelAttributeWithLogging(log, rel2, attCBasisOfOpinionRels, XL1.Sheets[sheet].Cells(row, 11).Text.Trim());

                                    itemShared.Relationship_AddItem(item); // no direction

                                    _sharedControlDictionary.Add(item.Id, itemShared);
                                }
                                RemoveRelFromList(list, rel);

                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                                extId = _causeControlsId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");
                            }
                        }
                        // cause-action
                        for (int row = 48; row <= 59; row++)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 14).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                GetItemNameAndDescription(text, out name, out desc);
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 13).Text.Trim());
                                extId = _causeControlActionsId + $"{order:D2}";

                                if (verbose) log.Log($"Processing '{extId}'");

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catCauseAction;

                                SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 20).Text);
                                SetAttributeWithLogging(log, item, attActionOwner, XL1.Sheets[sheet].Cells(row, 21).Text.Trim());
                                SetAttributeWithLogging(log, item, attPriority, XL1.Sheets[sheet].Cells(row, 22).Text.Trim());
                                SetAttributeWithLogging(log, item, attBaseline, XL1.Sheets[sheet].Cells(row, 23).Text.Trim());
                                SetAttributeWithLogging(log, item, attRevision, XL1.Sheets[sheet].Cells(row, 24).Text.Trim());
                                SetAttributeWithLogging(log, item, attPercComplete, XL1.Sheets[sheet].Cells(row, 25).Text.Trim().Replace("%", ""));
                                SetAttributeWithLogging(log, item, attStatus, XL1.Sheets[sheet].Cells(row, 26).Text.Trim());

                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));
                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 13).Text.Trim());
                                extId = _causeControlActionsId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");
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

                                if (verbose) log.Log($"Processing '{extId}'");

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catConsequence;
                                SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 55).Text);
                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));
                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 40).Text.Trim());
                                extId = _consequenceId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");
                            }
                        }
                        // consequence-control
                        for (int row = 48; row <= 59; row++)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 33).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                GetItemNameAndDescription(text, out name, out desc);
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 32).Text.Trim());
                                extId = _consequenceControlsId + $"{order:D2}";
                                if (verbose) log.Log($"Processing '{extId}'");

                                var words = name.Split(' ');
                                Item itemShared = null;
                                if (words.Any())
                                {
                                    // a shared item exists in the control library
                                    itemShared = controlStory.Item_FindByExternalId(words[0]);
                                }

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catConsequenceControl;
                                // always try to set these 
                                SetAttributeWithLogging(log, item, attControlOwner, XL1.Sheets[sheet].Cells(row, 39).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlOpinion, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 40).Text.Trim()));
                                SetAttributeWithLogging(log, item, attBasisOfOpinion, XL1.Sheets[sheet].Cells(row, 41).Text.Trim());
                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));

                                var rel = item.Relationship_AddItem(risk, "", Relationship.RelationshipDirection.BtoA);
                                SetRelAttributeWithLogging(log, rel, attControlOwnerRels, XL1.Sheets[sheet].Cells(row, 39).Text.Trim());
                                SetRelAttributeWithLogging(log, rel, attControlOpinionRels, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 40).Text.Trim()));
                                SetRelAttributeWithLogging(log, rel, attBasisOfOpinionRels, XL1.Sheets[sheet].Cells(row, 41).Text.Trim());

                                RemoveRelFromList(list, rel);

                                if (itemShared != null) // should be a shared control 
                                {
                                    log.Log($"DETECTED SHARED CONTROL '{words[0]}'");
                                    log.Log($"'{itemShared.Name}'");

                                    var rel2 = itemShared.Relationship_AddItem(risk, "RISK", Relationship.RelationshipDirection.BtoA);
                                    SetRelAttributeWithLogging(log, rel2, attCControlOwnerRels, XL1.Sheets[sheet].Cells(row, 39).Text.Trim());
                                    SetRelAttributeWithLogging(log, rel2, attCControlOpinionRels, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 40).Text.Trim()));
                                    SetRelAttributeWithLogging(log, rel2, attCBasisOfOpinionRels, XL1.Sheets[sheet].Cells(row, 41).Text.Trim());

                                    itemShared.Relationship_AddItem(item); // no direction

                                    _sharedControlDictionary.Add(item.Id, itemShared);
                                }
                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 32).Text.Trim());
                                extId = _consequenceControlsId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");

                            }
                        }
                        // consequence-action
                        for (int row = 48; row <= 59; row++)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 45).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                GetItemNameAndDescription(text, out name, out desc);
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 44).Text.Trim());
                                extId = _consequenceControlActionsId + $"{order:D2}";
                                if (verbose) log.Log($"Processing '{extId}'");

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catConsequenceAction;

                                SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 51).Text);
                                SetAttributeWithLogging(log, item, attActionOwner, XL1.Sheets[sheet].Cells(row, 52).Text.Trim());
                                SetAttributeWithLogging(log, item, attPriority, XL1.Sheets[sheet].Cells(row, 53).Text.Trim());
                                SetAttributeWithLogging(log, item, attBaseline, XL1.Sheets[sheet].Cells(row, 54).Text.Trim());
                                SetAttributeWithLogging(log, item, attRevision, XL1.Sheets[sheet].Cells(row, 55).Text.Trim());
                                SetAttributeWithLogging(log, item, attPercComplete, XL1.Sheets[sheet].Cells(row, 56).Text.Trim().Replace("%", ""));
                                SetAttributeWithLogging(log, item, attStatus, XL1.Sheets[sheet].Cells(row, 57).Text.Trim());
                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order - 36));
                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 44).Text.Trim());
                                extId = _consequenceControlActionsId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");

                            }
                        }

                        // Early Warning indicators
                        for (int row = 9; row <= 17; row += 2)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 22).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                GetItemNameAndDescription(text, out name, out desc);
                                order = counterEWI++;
                                extId = _ewiId + $"{order:D2}";
                                if (verbose) log.Log($"Processing '{extId}'");


                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catEWI;

                                SetAttributeWithLogging(log, item, attLinkedControlsTypes, XL1.Sheets[sheet].Cells(row, 19).Text);
                                SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 21).Text);
                                //SetAttributeWithLogging(log, item, attPrior, XL1.Sheets[sheet].Cells(row, 35).Text);
                                SetAttributeWithLogging(log, item, attCurrent, XL1.Sheets[sheet].Cells(row, 33).Text);
                                SetAttributeWithLogging(log, item, attTolerance, XL1.Sheets[sheet].Cells(row, 35).Text);
                                SetAttributeWithLogging(log, item, attWithinTolerance, XL1.Sheets[sheet].Cells(row, 37).Text);
                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));

                                var link = XL1.Sheets[sheet].Cells(row, 19).Text;
                            }
                            else if (deleteItems)
                            {
                                order = counterEWI++;
                                extId = _ewiId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");

                            }
                        }
                    }

                    if (verbose) log.Log($"Processing relationships");

                    // process the relationships for the Causes & Conseqences
                    // the template only allows for a max of 30 of each
                    for (int c = 1; c <= 30; c++)
                    {
                        // Causes
                        extId = _causeId + $"{c:D2}";
                        item = story.Item_FindByExternalId(extId);
                        if (item != null)
                        {
                            //log.Log($"processing rels for {item.Name}");
                            var rels = item.GetAttributeValueAsText(attLinkedControls);
                            if (verbose) log.Log($"Processing '{extId}' '{rels}'");
                            foreach (var r in rels.Split(','))
                            {
                                if (string.IsNullOrEmpty(r))
                                {
                                    log.Log($"Warning: '{item.Name}' has no related controls");
                                    break;
                                }
                                var i = GetInt(r);
                                var ex = _causeControlsId + $"{i:D2}";
                                var itm = story.Item_FindByExternalId(ex);
                                if (itm != null)
                                {
                                    RemoveRelFromList(list, item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.AtoB));
                                    if (_sharedControlDictionary.ContainsKey(itm.Id))
                                        _sharedControlDictionary[itm.Id].Relationship_AddItem(item, "", Relationship.RelationshipDirection.BtoA);
                                }
                                else
                                {
                                    log.Log($"Warning: Could not create relationships from '{item.Name}' to '{ex}'");
                                }
                            }
                        }

                        // Consequences
                        extId = _consequenceId + $"{c:D2}";
                        item = story.Item_FindByExternalId(extId);
                        if (item != null)
                        {
                            //log.Log($"processing rels for {item.Name}");
                            var rels = item.GetAttributeValueAsText(attLinkedControls);
                            if (verbose) log.Log($"Processing '{extId}' '{rels}'");
                            foreach (var r in rels.Split(','))
                            {
                                var i = GetInt(r);
                                var ex = _consequenceControlsId + $"{i:D2}";
                                var itm = story.Item_FindByExternalId(ex);
                                if (itm != null)
                                {
                                    RemoveRelFromList(list, item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.BtoA));
                                    if (_sharedControlDictionary.ContainsKey(itm.Id))
                                        _sharedControlDictionary[itm.Id].Relationship_AddItem(item, "", Relationship.RelationshipDirection.AtoB);
                                }
                                else
                                {
                                    log.Log($"Warning: Could not create relationships from '{item.Name}' to '{ex}'");
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
                            if (verbose) log.Log($"Processing '{extId}' '{rels}'");
                            foreach (var r in rels.Split(','))
                            {
                                if (!string.IsNullOrEmpty(r))
                                {
                                    var i = GetInt(r);
                                    var ex = _causeId + $"{i:D2}";
                                    if (relType == "Conseq.")
                                        ex = _consequenceId + $"{i:D2}";
                                    var itm = story.Item_FindByExternalId(ex);
                                    if (itm != null)
                                    {
                                        RemoveRelFromList(list, item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.AtoB));
                                        if (_sharedControlDictionary.ContainsKey(itm.Id))
                                            _sharedControlDictionary[itm.Id].Relationship_AddItem(item, "", Relationship.RelationshipDirection.BtoA);
                                    }
                                    else
                                    {
                                        log.Log($"Warning: Could not create relationships from '{item.Name}' to '{ex}'");
                                    }
                                }
                            }
                        }
                    }

                    // cause and consequence actions
                    for (int c = 1; c <= 36; c++)
                    {
                        // Cause-Actions
                        extId = _causeControlActionsId + $"{c:D2}";
                        item = story.Item_FindByExternalId(extId);
                        if (item != null)
                        {
                            //log.Log($"processing rels for {item.Name}");
                            var rels = item.GetAttributeValueAsText(attLinkedControls);
                            if (verbose) log.Log($"Processing '{extId}' '{rels}'");
                            foreach (var r in rels.Split(','))
                            {
                                var i = GetInt(r);
                                var ex = _causeControlsId + $"{i:D2}";
                                var itm = story.Item_FindByExternalId(ex);
                                if (itm != null)
                                {
                                    RemoveRelFromList(list, item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.AtoB));
                                    if (_sharedControlDictionary.ContainsKey(itm.Id))
                                        _sharedControlDictionary[itm.Id].Relationship_AddItem(item, "", Relationship.RelationshipDirection.BtoA);
                                }
                                else
                                {
                                    log.Log($"Warning: Could not create relationships from '{item.Name}' to '{ex}'");
                                }
                            }
                        }

                        // Consequence-Actions
                        extId = _consequenceControlActionsId + $"{c + 36:D2}";
                        item = story.Item_FindByExternalId(extId);
                        if (item != null)
                        {
                            //log.Log($"processing rels for {item.Name}");
                            var rels = item.GetAttributeValueAsText(attLinkedControls);
                            if (verbose) log.Log($"Processing '{extId}' '{rels}'");
                            foreach (var r in rels.Split(','))
                            {
                                var i = GetInt(r);
                                var ex = _consequenceControlsId + $"{i:D2}";
                                var itm = story.Item_FindByExternalId(ex);
                                if (itm != null)
                                {
                                    RemoveRelFromList(list, item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.AtoB));
                                    if (_sharedControlDictionary.ContainsKey(itm.Id))
                                        _sharedControlDictionary[itm.Id].Relationship_AddItem(item, "", Relationship.RelationshipDirection.BtoA);
                                }
                                else
                                {
                                    log.Log($"Warning: Could not create relationships from '{item.Name}' to '{ex}'");
                                }
                            }
                        }
                    }

                    if (deleteRels)
                    {
                        if (verbose) log.Log($"Deleting relationships");

                        foreach (var l in list)
                            story.Relationship_DeleteById(l.Key);
                    }

                    // process shared controls
                    foreach (var si in _sharedControlDictionary)
                    {
                        var li = story.Item_FindById(si.Key);
                        var sc = si.Value;

                        // add a resource link from the contol library to the control
                        var res = sc.Resource_FindByName(li.Name) ?? sc.Resource_AddName(li.Name);
                        res.Description = li.Story.Name;
                        res.Url = new Uri(li.Url);

                        // add a resource link from the contol library to the control
                        res = li.Resource_FindByName("Shared Control") ?? li.Resource_AddName("Shared Control");
                        res.Description = sc.Name;
                        res.Url = new Uri(sc.Url);
                    }

                    // add the directorate to all items
                    foreach (var i in story.Items)
                    {
                        SetAttributeWithLogging(log, i, attDirectorate, directorate);
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
                catch (Exception ex)
                {
                    log.LogError(ex);
                }
            }
            else
            {

                try
                {
                    var sheetNames = new string[] { "ERR", "ERR Cont Sheet 1", "ERR Cont Sheet 2" };

                    var XL1 = new Application();
                    var pathMlstn = XLFilename;
                    log.Log($"Opening Excel Doc " + pathMlstn);
                    var wbBowTie = XL1.Workbooks.Open(pathMlstn);
                    var sheet = sheetNames[0];

                    // validate template is correct version
                    //var version = XL1.Sheets["Version Control"].Cells[1, 26].Text;
                    //if (version != "SCApproved")
                    //{
                    //    log.Log($"Spreadsheet is not in the approved version, missing 'SCApproved' at Z1 in 'Version Control' ");
                    //    KillProcessByMainWindowHwnd(XL1.Application.Hwnd);
                    //    return;
                    //}
                    string version2 = XL1.Sheets["ERR"].Cells[1, 15].Text;
                    if (!version2.Contains("v5"))
                    {
                        log.Log($"Spreadsheet is not in the approved version, missing 'v45x' at 'O:1' in 'ERR' ");
                        KillProcessByMainWindowHwnd(XL1.Application.Hwnd);
                        return;
                    }

                    var list = new Dictionary<string, Relationship>();
                    foreach (var rel in story.Relationships)
                        list.Add(rel.Id, rel);

                    // set the story name
                    var level = XL1.Sheets[sheet].Cells(4, 4).Text;
                    var directorate = XL1.Sheets[sheet].Cells(5, 4).Text;
                    var subDirectorate = XL1.Sheets[sheet].Cells(6, 4).Text;

                    var title = XL1.Sheets[sheet].Cells(7, 4).Text;

                    story.Name = $"L{level}_{GetShortenedDirectorate(directorate)}_{title}";

                    Item risk = story.Item_FindByExternalId(_riskId) ?? story.Item_AddNew(title, false);
                    risk.ExternalId = _riskId;
                    risk.Description = XL1.Sheets[sheet].Cells(3, 20).Text;


                    risk.Category = catRisk;
                    SetAttributeWithLogging(log, risk, attClassification, XL1.Sheets[sheet].Cells(3, 4).Text);
                    SetAttributeWithLogging(log, risk, attRiskLevel, level);
                    SetAttributeWithLogging(log, risk, attDirectorate, directorate);
                   // SetAttributeWithLogging(log, risk, attrSubDirectorate, subDirectorate); not in V5 info spreadsheet.

                    SetAttributeWithLogging(log, risk, attRiskOwner, XL1.Sheets[sheet].Cells(8, 4).Text);
                    
                    SetAttributeWithLogging(log, risk, attManager, XL1.Sheets[sheet].Cells(9, 4).Text);

                //    SetAttributeWithLogging(log, risk, attImapactedArea, LookupRiskLabel(XL1.Sheets[sheet].Cells(10, 4).Text)); //TODO doesnt look right - ANW taken out as not in Spreadsheet
                    SetAttributeWithLogging(log, risk, attrRiskCategory, XL1.Sheets[sheet].Cells(11, 4).Text); //new ERR Cat

                    SetAttributeWithLogging(log, risk, attControlRating, XL1.Sheets[sheet].Cells(45, 36).Text); // new position on template
                    SetAttributeWithLogging(log, risk, attVersion, XL1.Sheets[sheet].Cells(12, 4).Text);
                    //SetAttributeWithLogging(log, risk, attLastUpdate, XL1.Sheets[sheet].Cells(13, 4).Text);
                    SetAttributeWithLogging(log, risk, attrRiskExposureFrame, XL1.Sheets[sheet].Cells(16, 4).Text);

                    // gross
                    SetAttributeWithLogging(log, risk, attGrossImpact, LookupRiskLabel(XL1.Sheets[sheet].Cells(22, 25).Text));
                    SetAttributeWithLogging(log, risk, attGrossLikelihood, LookupRiskLabel(XL1.Sheets[sheet].Cells(22, 28).Text));
                    SetAttributeWithLogging(log, risk, attGrossFinance, LookupRiskLabel(XL1.Sheets[sheet].Cells(22, 31).Text));
                    SetAttributeWithLogging(log, risk, attGrossRating, XL1.Sheets[sheet].Cells(22, 34).Text);
                    // target
                    SetAttributeWithLogging(log, risk, attTargetImpact, LookupRiskLabel(XL1.Sheets[sheet].Cells(23, 25).Text));
                    SetAttributeWithLogging(log, risk, attTargetLikelihood, LookupRiskLabel(XL1.Sheets[sheet].Cells(23, 28).Text));
                    SetAttributeWithLogging(log, risk, attTargetFinance, LookupRiskLabel(XL1.Sheets[sheet].Cells(23, 31).Text));
                    SetAttributeWithLogging(log, risk, attTargetRating, XL1.Sheets[sheet].Cells(23, 34).Text);

                    SetAttributeWithLogging(log, risk, attLikelihoodSafety, LookupRiskLabel(XL1.Sheets[sheet].Cells(26, 38).Text));
                    SetAttributeWithLogging(log, risk, attImpactSafety, LookupRiskLabel(XL1.Sheets[sheet].Cells(26, 36).Text));
                    SetAttributeWithLogging(log, risk, attAppetiteSafety, LookupYesNoRiskLabel(XL1.Sheets[sheet].Cells(28, 36).Text));
                    SetAttributeWithLogging(log, risk, attRationaleSafety, XL1.Sheets[sheet].Cells(26, 20).Text);

                    SetAttributeWithLogging(log, risk, attLikelihoodPerformance, LookupRiskLabel(XL1.Sheets[sheet].Cells(31, 38).Text));
                    SetAttributeWithLogging(log, risk, attImpactPerformance, LookupRiskLabel(XL1.Sheets[sheet].Cells(31, 36).Text));
                    SetAttributeWithLogging(log, risk, attAppetitePerformace, LookupYesNoRiskLabel(XL1.Sheets[sheet].Cells(33, 36).Text));
                    SetAttributeWithLogging(log, risk, attRationalePerformance, XL1.Sheets[sheet].Cells(31, 20).Text);

                    SetAttributeWithLogging(log, risk, attLikelihoodValue, LookupRiskLabel(XL1.Sheets[sheet].Cells(36, 38).Text));
                    SetAttributeWithLogging(log, risk, attImpactValue, LookupRiskLabel(XL1.Sheets[sheet].Cells(36, 36).Text));
                    SetAttributeWithLogging(log, risk, attAppetiteValue, LookupYesNoRiskLabel(XL1.Sheets[sheet].Cells(38, 36).Text));
                    SetAttributeWithLogging(log, risk, attRationaleValue, XL1.Sheets[sheet].Cells(36, 20).Text);

                    SetAttributeWithLogging(log, risk, attLikelihoodPolitical, LookupRiskLabel(XL1.Sheets[sheet].Cells(41, 38).Text));
                    SetAttributeWithLogging(log, risk, attImpactPolitical, LookupRiskLabel(XL1.Sheets[sheet].Cells(41, 36).Text));
                    SetAttributeWithLogging(log, risk, attAppetitePolitical, LookupYesNoRiskLabel(XL1.Sheets[sheet].Cells(43, 36).Text));
                    SetAttributeWithLogging(log, risk, attRationalePolitical, XL1.Sheets[sheet].Cells(41, 20).Text);

                    SetAttributeWithLogging(log, risk, attLikelihood, LookupRiskLabel(XL1.Sheets[sheet].Cells(2, 80).Text)); //TODO?
                    SetAttributeWithLogging(log, risk, attImpact, LookupRiskLabel(XL1.Sheets[sheet].Cells(1, 80).Text)); //TODO?
                    SetAttributeWithLogging(log, risk, attRationale, XL1.Sheets[sheet].Cells(46, 20).Text);

                    //SetAttributeWithLogging(log, risk, attReportingPriority, GetReportingPriority(order));

                    // Risk Trajectory


               

                    SetAttributeWithLogging(log, risk, attRiskM1Impact, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(29, 3).Text));
                    SetAttributeWithLogging(log, risk, attRiskM1Likelihood, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(29, 4).Text));
                    SetAttributeWithLogging(log, risk, attRiskM1Finance, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(29, 5).Text));
                    SetAttributeWithLogging(log, risk, attRiskM1Rating, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(29, 6).Text));

                    SetAttributeWithLogging(log, risk, attRiskM2Impact, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(30, 3).Text));
                    SetAttributeWithLogging(log, risk, attRiskM2Likelihood, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(30, 4).Text));
                    SetAttributeWithLogging(log, risk, attRiskM2Finance, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(30, 5).Text));
                    SetAttributeWithLogging(log, risk, attRiskM2Rating, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(30, 6).Text));

                    SetAttributeWithLogging(log, risk, attRiskM3Impact, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(31, 3).Text));
                    SetAttributeWithLogging(log, risk, attRiskM3Likelihood, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(31, 4).Text));
                    SetAttributeWithLogging(log, risk, attRiskM3Finance, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(31, 5).Text));
                    SetAttributeWithLogging(log, risk, attRiskM3Rating, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(31, 6).Text));

                    SetAttributeWithLogging(log, risk, attRiskM4Impact, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(32, 3).Text));
                    SetAttributeWithLogging(log, risk, attRiskM4Likelihood, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(32, 4).Text));
                    SetAttributeWithLogging(log, risk, attRiskM4Finance, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(32, 5).Text));
                    SetAttributeWithLogging(log, risk, attRiskM4Rating, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(32, 6).Text));

                    SetAttributeWithLogging(log, risk, attRiskM5Impact, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(33, 3).Text));
                    SetAttributeWithLogging(log, risk, attRiskM5Likelihood, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(33, 4).Text));
                    SetAttributeWithLogging(log, risk, attRiskM5Finance, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(33, 5).Text));
                    SetAttributeWithLogging(log, risk, attRiskM5Rating, LookupRiskLabel(XL1.Sheets["Risk Trajectory"].Cells(33, 6).Text));

                  
                    SetAttributeWithLogging(log, risk, attRiskGrossToNet, XL1.Sheets["Risk Trajectory"].Cells(8, 44).Text);
                    SetAttributeWithLogging(log, risk, attRiskNetToTarget, XL1.Sheets["Risk Trajectory"].Cells(19, 44).Text);
                    SetAttributeWithLogging(log, risk, attRiskType, XL1.Sheets["ERR"].Cells(2, 20).Text); // THREAT or OPPORTUNITY



                    string tagText = XL1.Sheets["Version Control"].Cells[4, 8].Text; //TODO - ask DAVID not on new sheet - found on version h4 - needs splitting out i think.
                    var tags = tagText.Split(',');
                    foreach (var t in tags)
                    {

                        SetTagWithLogging(log,story, risk, t);
                        
                    }


                    //add SHE tag
                    SetTagWithLogging(log, story, risk, XL1.Sheets["ERR"].Cells(25, 28).Text);
                    //add Performance tag
                    SetTagWithLogging(log, story, risk, XL1.Sheets["ERR"].Cells(30, 28).Text);
                   //add Performance tag
                    SetTagWithLogging(log, story, risk, XL1.Sheets["ERR"].Cells(35, 28).Text);
                   //add Performance tag
                    SetTagWithLogging(log, story, risk, XL1.Sheets["ERR"].Cells(40, 28).Text);
                    


                    Item item;
                    string extId;
                    int order;
                    int counterEWI = 1;
                    string name;
                    string desc;
                    string text;
                    int sht;

                    //special code to delete old consequence actions
                    if (deleteItems)
                    {
                        for (int i = 1; i <= 36; i++)
                        {
                            extId = _consequenceControlActionsId + $"{i:D2}";
                            DeleteItemWithLogging(log, story, extId);
                        }

                        DeleteItemWithLogging(log, story, _causeControls + $"{0:D2}");
                        DeleteItemWithLogging(log, story, _consequenceControls + $"{0:D2}");
                    }


                    // data can be in the same place on 3 sheets (continuation sheets)
                    for (sht = 0; sht < 3; sht++)
                    {
                        sheet = sheetNames[sht];

                        try
                        {
                            var testsheet = XL1.Sheets[sheet];
                            if (testsheet == null)
                            {
                                log.Log($"Sheet '{sheet}' does not exist, skipping.");
                                continue;
                            }
                        }
                        catch (Exception exsheet)
                        {
                            log.Log($"Sheet '{sheet}' does not exist, skipping.");
                            continue;
                        }


                        if (verbose) log.Log($"Processing Sheet{sheet}");

                        // cause
                        for (int row = 25; row < 43; row += 2)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 3).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {

                                GetItemNameAndDescription(text, out name, out desc);
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                                extId = _causeId + $"{order:D2}";
                                if (verbose) log.Log($"Processing '{extId}'");

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catCause;
                                SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 17).Text);
                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));
                                SetAttributeWithLogging(log, item, attControlIndustry, XL1.Sheets[sheet].Cells(row, 16).Text); //anw added
                                //add tag as well
                                SetTagWithLogging(log,story,item, XL1.Sheets[sheet].Cells(row, 16).Text);
                               
                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                                extId = _causeId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting {extId}");
                            }
                        }
                        // cause-control
                        for (int row = 53; row <= 119; row+=6)
                        {

                            text = XL1.Sheets[sheet].Cells(row, 4).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                GetItemNameAndDescription(text, out name, out desc);
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                                extId = _causeControlsId + $"{order:D2}";
                                if (verbose) log.Log($"Processing '{extId}'");

                                var words = name.Split(' ');
                                Item itemShared = null;
                                if (words.Any())
                                {
                                    // a shared item exists in the control library
                                    itemShared = controlStory.Item_FindByExternalId(words[0]);
                                }

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catCauseControl;
                                // awlays try to set these
                                SetAttributeWithLogging(log, item, attControlOwner, XL1.Sheets[sheet].Cells(row, 10).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlOpinion, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 11).Text.Trim()));
                                SetAttributeWithLogging(log, item, attBasisOfOpinion, XL1.Sheets[sheet].Cells(row, 12).Text.Trim());
                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));

                                //anw added 
                                SetAttributeWithLogging(log, item, attControlOpinionSource, XL1.Sheets[sheet].Cells(row+2, 11).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlOpinionRationale, XL1.Sheets[sheet].Cells(row+5, 3).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlIndustry, XL1.Sheets[sheet].Cells(row, 3).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlIndustryPartner, XL1.Sheets[sheet].Cells((row + 1), 3).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlIndustryVisibility, XL1.Sheets[sheet].Cells((row + 3), 3).Text.Trim());
                                SetTagWithLogging(log, story,item, XL1.Sheets[sheet].Cells(row, 3).Text.Trim());

                                var rel = item.Relationship_AddItem(risk, "", Relationship.RelationshipDirection.AtoB);
                                SetRelAttributeWithLogging(log, rel, attControlOwnerRels, XL1.Sheets[sheet].Cells(row, 10).Text.Trim());
                                SetRelAttributeWithLogging(log, rel, attControlOpinionRels, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 11).Text.Trim()));
                                SetRelAttributeWithLogging(log, rel, attBasisOfOpinionRels, XL1.Sheets[sheet].Cells(row, 12).Text.Trim());

                                SetRelAttributeWithLogging(log, rel, attOpinionRationalerRels, XL1.Sheets[sheet].Cells(row+5, 3).Text.Trim());


                                if (itemShared != null) // should be a shared control 
                                {
                                    log.Log($"DETECTED SHARED CONTROL '{words[0]}'");
                                    log.Log($"'{itemShared.Name}'");

                                    var rel2 = itemShared.Relationship_AddItem(risk, "RISK", Relationship.RelationshipDirection.AtoB);
                                    SetRelAttributeWithLogging(log, rel2, attCControlOwnerRels, XL1.Sheets[sheet].Cells(row, 10).Text.Trim());
                                    SetRelAttributeWithLogging(log, rel2, attCControlOpinionRels, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 11).Text.Trim()));
                                    SetRelAttributeWithLogging(log, rel2, attCBasisOfOpinionRels, XL1.Sheets[sheet].Cells(row, 12).Text.Trim());
                                    SetRelAttributeWithLogging(log, rel2, attOpinionRationalerRels, XL1.Sheets[sheet].Cells(row + 5, 3).Text.Trim());

                                    itemShared.Relationship_AddItem(item); // no direction

                                    _sharedControlDictionary.Add(item.Id, itemShared);
                                }
                                RemoveRelFromList(list, rel);

                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 2).Text.Trim());
                                extId = _causeControlsId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");
                            }
                        }
                        // cause-action
                        for (int row = 53; row <= 119; row+=6)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 15).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                GetItemNameAndDescription(text, out name, out desc);
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 14).Text.Trim());
                                extId = _causeControlActionsId + $"{order:D2}";

                                if (verbose) log.Log($"Processing '{extId}'");

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catCauseAction;

                                SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 21).Text);
                                SetAttributeWithLogging(log, item, attActionOwner, XL1.Sheets[sheet].Cells(row, 22).Text.Trim());
                                SetAttributeWithLogging(log, item, attPriority, XL1.Sheets[sheet].Cells(row, 23).Text.Trim());
                                SetAttributeWithLogging(log, item, attBaseline, XL1.Sheets[sheet].Cells(row, 24).Text.Trim());
                                SetAttributeWithLogging(log, item, attRevision, XL1.Sheets[sheet].Cells(row, 25).Text.Trim());
                                SetAttributeWithLogging(log, item, attPercComplete, XL1.Sheets[sheet].Cells(row, 26).Text.Trim().Replace("%", ""));
                                SetAttributeWithLogging(log, item, attStatus, XL1.Sheets[sheet].Cells(row, 27).Text.Trim());


                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));
                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 14).Text.Trim());
                                extId = _causeControlActionsId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");
                            }
                        }

                        // consequences
                        for (int row = 23; row < 43; row += 2)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 42).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                GetItemNameAndDescription(text, out name, out desc);
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 41).Text.Trim());
                                extId = _consequenceId + $"{order:D2}";

                                if (verbose) log.Log($"Processing '{extId}'");

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catConsequence;
                                SetAttributeWithLogging(log, item, attControlIndustry, XL1.Sheets[sheet].Cells(row, 55).Text); //anw added
                                SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 57).Text);
                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));
                                SetTagWithLogging(log, story,item,XL1.Sheets[sheet].Cells(row, 55).Text.Trim());
                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 41).Text.Trim());
                                extId = _consequenceId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");
                            }
                        }
                        // consequence-control
                        for (int row = 53; row <= 119; row+=6)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 35).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                GetItemNameAndDescription(text, out name, out desc);
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 33).Text.Trim());
                                extId = _consequenceControlsId + $"{order:D2}";
                                if (verbose) log.Log($"Processing '{extId}'");

                                var words = name.Split(' ');
                                Item itemShared = null;
                                if (words.Any())
                                {
                                    // a shared item exists in the control library
                                    itemShared = controlStory.Item_FindByExternalId(words[0]);
                                }

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catConsequenceControl;
                                // always try to set these 
                                SetAttributeWithLogging(log, item, attControlOwner, XL1.Sheets[sheet].Cells(row, 40).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlOpinion, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 41).Text.Trim()));
                                SetAttributeWithLogging(log, item, attBasisOfOpinion, XL1.Sheets[sheet].Cells(row, 42).Text.Trim());

                                //anw added 
                                SetAttributeWithLogging(log, item, attControlOpinionSource, XL1.Sheets[sheet].Cells(row + 2, 42).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlOpinionRationale, XL1.Sheets[sheet].Cells(row + 5, 34).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlIndustry, XL1.Sheets[sheet].Cells(row, 34).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlIndustryPartner, XL1.Sheets[sheet].Cells((row + 1), 34).Text.Trim());
                                SetAttributeWithLogging(log, item, attControlIndustryVisibility, XL1.Sheets[sheet].Cells((row + 3), 34).Text.Trim());
                                SetTagWithLogging(log, story, item, XL1.Sheets[sheet].Cells(row, 34).Text.Trim());


                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));

                                var rel = item.Relationship_AddItem(risk, "", Relationship.RelationshipDirection.BtoA);
                                SetRelAttributeWithLogging(log, rel, attControlOwnerRels, XL1.Sheets[sheet].Cells(row, 40).Text.Trim());
                                SetRelAttributeWithLogging(log, rel, attControlOpinionRels, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 41).Text.Trim()));
                                SetRelAttributeWithLogging(log, rel, attBasisOfOpinionRels, XL1.Sheets[sheet].Cells(row, 42).Text.Trim());
                                SetRelAttributeWithLogging(log, rel, attOpinionRationalerRels, XL1.Sheets[sheet].Cells(row + 5, 34).Text.Trim());

                                RemoveRelFromList(list, rel);

                                if (itemShared != null) // should be a shared control 
                                {
                                    log.Log($"DETECTED SHARED CONTROL '{words[0]}'");
                                    log.Log($"'{itemShared.Name}'");

                                    var rel2 = itemShared.Relationship_AddItem(risk, "RISK", Relationship.RelationshipDirection.BtoA);
                                    SetRelAttributeWithLogging(log, rel2, attCControlOwnerRels, XL1.Sheets[sheet].Cells(row, 40).Text.Trim());
                                    SetRelAttributeWithLogging(log, rel2, attCControlOpinionRels, LookupControlOpinion(XL1.Sheets[sheet].Cells(row, 41).Text.Trim()));
                                    SetRelAttributeWithLogging(log, rel2, attCBasisOfOpinionRels, XL1.Sheets[sheet].Cells(row, 42).Text.Trim());
                                    SetRelAttributeWithLogging(log, rel2, attOpinionRationalerRels, XL1.Sheets[sheet].Cells(row + 5, 34).Text.Trim());

                                    itemShared.Relationship_AddItem(item); // no direction

                                    _sharedControlDictionary.Add(item.Id, itemShared);
                                }
                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 33).Text.Trim());
                                extId = _consequenceControlsId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");

                            }
                        }
                        // consequence-action
                        for (int row = 53; row <= 119; row+=6)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 46).Text;
                                                                                                                                                                                                                                                                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                GetItemNameAndDescription(text, out name, out desc);
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 45).Text.Trim());
                                extId = _consequenceControlActionsId + $"{order:D2}";
                                if (verbose) log.Log($"Processing '{extId}'");

                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catConsequenceAction;

                                SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 52).Text);
                                SetAttributeWithLogging(log, item, attActionOwner, XL1.Sheets[sheet].Cells(row, 53).Text.Trim());
                                SetAttributeWithLogging(log, item, attPriority, XL1.Sheets[sheet].Cells(row, 54).Text.Trim());
                                SetAttributeWithLogging(log, item, attBaseline, XL1.Sheets[sheet].Cells(row, 55).Text.Trim());
                                SetAttributeWithLogging(log, item, attRevision, XL1.Sheets[sheet].Cells(row, 56).Text.Trim());
                                SetAttributeWithLogging(log, item, attPercComplete, XL1.Sheets[sheet].Cells(row, 57).Text.Trim().Replace("%", ""));
                                SetAttributeWithLogging(log, item, attStatus, XL1.Sheets[sheet].Cells(row, 58).Text.Trim());
                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order - 36));
                            }
                            else if (deleteItems)
                            {
                                order = GetInt(XL1.Sheets[sheet].Cells(row, 45).Text.Trim());
                                extId = _consequenceControlActionsId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");

                            }
                        }

                        // Early Warning indicators
                        int EWIChartCount = 0;
                        for (int row = 9; row <= 17; row += 2)
                        {
                            text = XL1.Sheets[sheet].Cells(row, 23).Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                GetItemNameAndDescription(text, out name, out desc);
                                order = counterEWI++;
                                extId = _ewiId + $"{order:D2}";
                                if (verbose) log.Log($"Processing '{extId}'");


                                item = story.Item_FindByExternalId(extId) ?? story.Item_AddNew(name, false);
                                item.ExternalId = extId;
                                item.Name = name;
                                item.Description = desc;
                                item.Category = catEWI;

                                SetAttributeWithLogging(log, item, attLinkedControlsTypes, XL1.Sheets[sheet].Cells(row, 20).Text);
                                SetAttributeWithLogging(log, item, attLinkedControls, XL1.Sheets[sheet].Cells(row, 22).Text);
                                //SetAttributeWithLogging(log, item, attPrior, XL1.Sheets[sheet].Cells(row, 35).Text);
                                SetAttributeWithLogging(log, item, attCurrent, XL1.Sheets[sheet].Cells(row, 34).Text);
                                SetAttributeWithLogging(log, item, attTolerance, XL1.Sheets[sheet].Cells(row, 36).Text);
                                SetAttributeWithLogging(log, item, attWithinTolerance, XL1.Sheets[sheet].Cells(row, 38).Text);
                                SetAttributeWithLogging(log, item, attSortOrder, order);
                                SetAttributeWithLogging(log, item, attReportingPriority, GetReportingPriority(order));

                                var link = XL1.Sheets[sheet].Cells(row, 20).Text;

                                //build EWI data for charts.

                                int EWIrow = row + 1;
                                EWIrow = 10 + EWIChartCount;
                               

                                var csvStringDelta = "";
                                var csvStringActual = "";
                                var csvStringTolerance = "";
                                for (int i = 91; i <= 103; i++)
                                {
                                    //delta
                                    var textEWI = XL1.Sheets["EWI"].Cells(EWIrow, i).Text;
                                    double val = GetDouble(textEWI);
                                    if (val > -99999)
                                    {
                                        csvStringDelta = csvStringDelta + "," + val.ToString();
                                    }
                                    //actual
                                    var textEWI2 = XL1.Sheets["EWI"].Cells(EWIrow+1, i).Text;
                                    double val2 = GetDouble(textEWI2);
                                    if (val2 > -99999)
                                    {
                                        csvStringActual = csvStringActual + "," + val2.ToString();
                                    }//tolersnce
                                    var textEWI3 = XL1.Sheets["EWI"].Cells(EWIrow+2, i).Text;
                                    double val3 = GetDouble(textEWI3);
                                    if (val3 > -99999)
                                    {
                                        csvStringTolerance = csvStringTolerance + "," + val3.ToString();
                                    }



                                }
                                EWIChartCount += 5;
                                if (csvStringDelta.Length > 0)
                                {
                                    csvStringDelta = "$$D" + csvStringDelta.Substring(1);
                                }
                                if (csvStringActual.Length > 0)
                                {
                                    csvStringActual = "$$A" + csvStringActual.Substring(1);
                                }
                                if (csvStringTolerance.Length > 0)
                                {
                                    csvStringTolerance = "$$T" + csvStringTolerance.Substring(1);
                                }
                                SetAttributeWithLogging(log, item, attEWISparkVals, csvStringDelta);
                                SetAttributeWithLogging(log, item, attEWISparkValsActual, csvStringActual);
                                SetAttributeWithLogging(log, item, attEWISparkValsTolerance, csvStringTolerance);

                            }
                            else if (deleteItems)
                            {
                                order = counterEWI++;
                                extId = _ewiId + $"{order:D2}";
                                DeleteItemWithLogging(log, story, extId);
                                if (verbose) log.Log($"Deleting '{extId}'");

                            }
                        }
                    }

                    if (verbose) log.Log($"Processing relationships");

                    // process the relationships for the Causes & Conseqences
                    // the template only allows for a max of 30 of each
                    for (int c = 1; c <= 30; c++)
                    {
                        // Causes
                        extId = _causeId + $"{c:D2}";
                        item = story.Item_FindByExternalId(extId);
                        if (item != null)
                        {
                            //log.Log($"processing rels for {item.Name}");
                            var rels = item.GetAttributeValueAsText(attLinkedControls);
                            if (verbose) log.Log($"Processing '{extId}' '{rels}'");
                            foreach (var r in rels.Split(','))
                            {
                                if (string.IsNullOrEmpty(r))
                                {
                                    log.Log($"Warning: '{item.Name}' has no related controls");
                                    break;
                                }
                                var i = GetInt(r);
                                var ex = _causeControlsId + $"{i:D2}";
                                var itm = story.Item_FindByExternalId(ex);
                                if (itm != null)
                                {
                                    RemoveRelFromList(list, item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.AtoB));
                                    if (_sharedControlDictionary.ContainsKey(itm.Id))
                                        _sharedControlDictionary[itm.Id].Relationship_AddItem(item, "", Relationship.RelationshipDirection.BtoA);
                                }
                                else
                                {
                                    log.Log($"Warning: Could not create relationships from '{item.Name}' to '{ex}'");
                                }
                            }
                        }

                        // Consequences
                        extId = _consequenceId + $"{c:D2}";
                        item = story.Item_FindByExternalId(extId);
                        if (item != null)
                        {
                            //log.Log($"processing rels for {item.Name}");
                            var rels = item.GetAttributeValueAsText(attLinkedControls);
                            if (verbose) log.Log($"Processing '{extId}' '{rels}'");
                            foreach (var r in rels.Split(','))
                            {
                                var i = GetInt(r);
                                var ex = _consequenceControlsId + $"{i:D2}";
                                var itm = story.Item_FindByExternalId(ex);
                                if (itm != null)
                                {
                                    RemoveRelFromList(list, item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.BtoA));
                                    if (_sharedControlDictionary.ContainsKey(itm.Id))
                                        _sharedControlDictionary[itm.Id].Relationship_AddItem(item, "", Relationship.RelationshipDirection.AtoB);
                                }
                                else
                                {
                                    log.Log($"Warning: Could not create relationships from '{item.Name}' to '{ex}'");
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
                            if (verbose) log.Log($"Processing '{extId}' '{rels}'");
                            foreach (var r in rels.Split(','))
                            {
                                if (!string.IsNullOrEmpty(r))
                                {
                                    var i = GetInt(r);
                                    var ex = _causeId + $"{i:D2}";
                                    if (relType == "Conseq.")
                                        ex = _consequenceId + $"{i:D2}";
                                    var itm = story.Item_FindByExternalId(ex);
                                    if (itm != null)
                                    {
                                        RemoveRelFromList(list, item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.AtoB));
                                        if (_sharedControlDictionary.ContainsKey(itm.Id))
                                            _sharedControlDictionary[itm.Id].Relationship_AddItem(item, "", Relationship.RelationshipDirection.BtoA);
                                    }
                                    else
                                    {
                                        log.Log($"Warning: Could not create relationships from '{item.Name}' to '{ex}'");
                                    }
                                }
                            }
                        }
                    }

                    // cause and consequence actions
                    for (int c = 1; c <= 36; c++)
                    {
                        // Cause-Actions
                        extId = _causeControlActionsId + $"{c:D2}";
                        item = story.Item_FindByExternalId(extId);
                        if (item != null)
                        {
                            //log.Log($"processing rels for {item.Name}");
                            var rels = item.GetAttributeValueAsText(attLinkedControls);
                            if (verbose) log.Log($"Processing '{extId}' '{rels}'");
                            foreach (var r in rels.Split(','))
                            {
                                var i = GetInt(r);
                                var ex = _causeControlsId + $"{i:D2}";
                                var itm = story.Item_FindByExternalId(ex);
                                if (itm != null)
                                {
                                    RemoveRelFromList(list, item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.AtoB));
                                    if (_sharedControlDictionary.ContainsKey(itm.Id))
                                        _sharedControlDictionary[itm.Id].Relationship_AddItem(item, "", Relationship.RelationshipDirection.BtoA);
                                }
                                else
                                {
                                    log.Log($"Warning: Could not create relationships from '{item.Name}' to '{ex}'");
                                }
                            }
                        }

                        // Consequence-Actions
                        extId = _consequenceControlActionsId + $"{c + 36:D2}";
                        item = story.Item_FindByExternalId(extId);
                        if (item != null)
                        {
                            //log.Log($"processing rels for {item.Name}");
                            var rels = item.GetAttributeValueAsText(attLinkedControls);
                            if (verbose) log.Log($"Processing '{extId}' '{rels}'");
                            foreach (var r in rels.Split(','))
                            {
                                var i = GetInt(r);
                                var ex = _consequenceControlsId + $"{i:D2}";
                                var itm = story.Item_FindByExternalId(ex);
                                if (itm != null)
                                {
                                    RemoveRelFromList(list, item.Relationship_AddItem(itm, "", Relationship.RelationshipDirection.AtoB));
                                    if (_sharedControlDictionary.ContainsKey(itm.Id))
                                        _sharedControlDictionary[itm.Id].Relationship_AddItem(item, "", Relationship.RelationshipDirection.BtoA);
                                }
                                else
                                {
                                    log.Log($"Warning: Could not create relationships from '{item.Name}' to '{ex}'");
                                }
                            }
                        }
                    }

                    if (deleteRels)
                    {
                        if (verbose) log.Log($"Deleting relationships");

                        foreach (var l in list)
                            story.Relationship_DeleteById(l.Key);
                    }

                    // process shared controls
                    foreach (var si in _sharedControlDictionary)
                    {
                        var li = story.Item_FindById(si.Key);
                        var sc = si.Value;

                        // add a resource link from the contol library to the control
                        var res = sc.Resource_FindByName(li.Name) ?? sc.Resource_AddName(li.Name);
                        res.Description = li.Story.Name;
                        res.Url = new Uri(li.Url);

                        // add a resource link from the contol library to the control
                        res = li.Resource_FindByName("Shared Control") ?? li.Resource_AddName("Shared Control");
                        res.Description = sc.Name;
                        res.Url = new Uri(sc.Url);
                    }

                    // add the directorate to all items
                    foreach (var i in story.Items)
                    {
                        SetAttributeWithLogging(log, i, attDirectorate, directorate);
                       
                        SetAttributeWithLogging(log, i, attRiskLevel, level);
                        SetTagWithLogging(log, story, i, XL1.Sheets["ERR"].Cells(6,4).Text.Trim());
                       

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
                catch (Exception ex)
                {
                    log.LogError(ex);
                }
            }
        }

        public static void UpdateRiskCountOnControlStory(Story controlStory, Logger logger)
        {
            EnsureControlStoryHasRightStructure(controlStory, logger);

            var attRiskCount = controlStory.Attribute_FindByName(_attrRelatedRisks);

            foreach (var i in controlStory.Items)
            {
                int count = 0;
                foreach (var r in i.Relationships)
                {
                    if (r.Comment == "RISK")
                        count++;
                }
                SetAttributeWithLogging(logger, i, attRiskCount, count);
            }


        }

        private static void RemoveRelFromList(Dictionary<string, Relationship> list, Relationship rel)
        {
            if (list.ContainsKey(rel.Id))
                list.Remove(rel.Id);
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

        private static void DeleteItemWithLogging(Logger log, Story story, string itemExtId)
        {
            var item = story.Item_FindByExternalId(itemExtId);
            if (item != null)
            {
                log.Log($"Deleting item '{itemExtId}'");
                story.Item_DeleteById(item.Id);
            }
        }

        private static void SetRelAttributeWithLogging(Logger log, Relationship rel, RelationshipAttribute att, string value)
        {
            try
            {
                rel.SetAttributeValue(att, value);
            }
            catch (Exception e)
            {
                var start = (rel.Item1 == null) ? rel.Item1.Name : $"unkown";
                var end = (rel.Item2==null)? rel.Item2.Name : $"unkown: ";

                log.Log($"ERROR unable to set relationship attribute='{att.Name}', value='{value}'");
                log.Log($"Relationship [{rel.Id}] between '{start}' and '{end}'");
                log.Log($"Error: {e.Message}");
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
                var b = true;
                if (value.Contains("%")) 
                {
                    value = value.Replace("%", "").Trim(); // important - must not get stuck in a loop
                    double dbl;
                    if (double.TryParse(value, out dbl))
                    {
                        b = false; // prevent the logging from the first instance
                        SetAttributeWithLogging(log, item, att, $"{dbl/100}");
                    }

                }
                if (b)
                    log.Log($"Error: {e.Message}, value='{value}', item='{item.Name}', attribute='{att.Name}'");
            }
            
        }

        private static void SetTagWithLogging(Logger log,Story story, Item item, string value)
        {
            try
            {
                Debug.WriteLine($"Setting Tag {value.Trim()} for {item.Name}");
                ItemTag tag = story.ItemTag_FindByName(value.Trim());
              
                    if (tag != null) {
                        if (item.Tag_FindById(tag.Id) == null) item.Tag_AddNew(tag);
                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(value))
                        {
                            item.Tag_AddNew(value.Trim());
                        }
                    }


            }
            catch (Exception e)
            {
                
                    log.Log($"Error: {e.Message}, value='{value}', item='{item.Name}'");
            }

        }

        

        private static int GetInt(string str)
        {
            int ret;
            if (Int32.TryParse(str, out ret))
                return ret;
            return 0;
        }

        private static double GetDouble(string str)
        {
            double ret;
            if (double.TryParse(str, out ret))
                return ret;
            return -99999;
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
                    CopyValues(sheet, XLS, row, 2, XLD);
                    CopyValues(sheet, XLS, row, 3, XLD);
                    CopyValues(sheet, XLS, row, 16, XLD);
                    //conseqeunces
                    CopyValues(sheet, XLS, row, 40, XLD);
                    CopyValues(sheet, XLS, row, 41, XLD);
                    CopyValues(sheet, XLS, row, 56, XLD);
                }

                for (int row = 48; row <= 59; row += 1)
                {
                    // cause controls
                    CopyValues(sheet, XLS, row, 2, XLD);
                    CopyValues(sheet, XLS, row, 3, XLD);
                    CopyValues(sheet, XLS, row, 9, XLD);
                    CopyValues(sheet, XLS, row, 10, XLD);
                    CopyValues(sheet, XLS, row, 11, XLD);

                    // cause action
                    CopyValues(sheet, XLS, row, 13, XLD);
                    CopyValues(sheet, XLS, row, 14, XLD);
                    CopyValues(sheet, XLS, row, 21, XLD);
                    CopyValues(sheet, XLS, row, 22, XLD);
                    CopyValues(sheet, XLS, row, 23, XLD);
                    CopyValues(sheet, XLS, row, 24, XLD);
                    CopyValues(sheet, XLS, row, 25, XLD);
                    CopyValues(sheet, XLS, row, 26, XLD);

                    // consequence controls
                    CopyValues(sheet, XLS, row, 2 + 30, XLD);
                    CopyValues(sheet, XLS, row, 3 + 30, XLD);
                    CopyValues(sheet, XLS, row, 9 + 30, XLD);
                    CopyValues(sheet, XLS, row, 10 + 30, XLD);
                    CopyValues(sheet, XLS, row, 11 + 30, XLD);

                    // conseuence actions
                    CopyValues(sheet, XLS, row, 13 + 31, XLD);
                    CopyValues(sheet, XLS, row, 14 + 31, XLD);
                    CopyValues(sheet, XLS, row, 21 + 31, XLD);
                    CopyValues(sheet, XLS, row, 22 + 31, XLD);
                    CopyValues(sheet, XLS, row, 23 + 31, XLD);
                    CopyValues(sheet, XLS, row, 24 + 31, XLD);
                    CopyValues(sheet, XLS, row, 25 + 31, XLD);
                    CopyValues(sheet, XLS, row, 26 + 31, XLD);
                }
            }

            // copy document control fields
            // find start row
            var vc = "Version Control";
            int rowDCF = -1;
            for (int r = 1; r < 100; r++)
            {
                string s = XLS.Sheets[vc].Cells[r, 1].Text;
                Debug.WriteLine(s);
                if (XLS.Sheets[vc].Cells[r, 1].Text.Trim() == "Document Version Control")
                {
                    rowDCF = r;
                    break;
                }
            }
            int rowD = 39;
            try
            {
                if (rowDCF != -1)
                {
                    int rowS = rowDCF + 3;
                    while (!string.IsNullOrEmpty(XLS.Sheets[vc].Cells[rowS, 1].Text))
                    {
                        for (int c = 1; c < 7; c++)
                            XLD.Sheets[vc].Cells[rowD, c] = XLS.Sheets[vc].Cells[rowS, c];
                        rowD++;
                        rowS++;
                    }
                }
                XLD.Sheets[vc].Cells[rowD, 1] = DateTime.Now.ToShortDateString();
                XLD.Sheets[vc].Cells[rowD, 2] = "2.0";
                XLD.Sheets[vc].Cells[rowD, 3] = "Automatic transfer of ERR to version 2.0 of the template";
                XLD.Sheets[vc].Cells[rowD, 4] = "All";
            }
            catch (Exception e)
            {
                log.LogError($"Bad data in spreadsheet - cannot copy accoss doc control");
            }
            try
            {
                GC.Collect();
                GC.WaitForFullGCComplete();
                wbSource.Close(false);
                Marshal.ReleaseComObject(wbSource);

                wbTemplate.SaveAs(newFile.Replace("xlsx", "xlsm"),
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);

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
            catch (Exception eBad)
            {
                log.LogError(eBad.Message);
            }
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
