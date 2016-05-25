using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.ComponentModel;
using Microsoft.SharePoint;
using System.Collections;
using Microsoft.SharePoint.Taxonomy;
using Microsoft.SharePoint.Navigation;


namespace Sharepoint.Helper
    public class TermStoreHelper
    {
        public static string idStore
        {
            get { return "00000000-0000-0000-C000-000000000046"; }
        }

        public static TermStore GetTermStore(SPSite site)
        {
            LogHelper logger = LogHelper.Instance;
            TaxonomySession session = new TaxonomySession(site);
            TermStore termStore = null;

            // get the TaxonomyField from the Site Columns in the sitecollectionTaxonomyField 
            //field = (TaxonomyField)site.RootWeb.Fields[TAXONOMYFIELDID];

            // get the Term Store ID from the field
            //Guid termStoreId = field.SspId;

            // Open a taxonomysession and get the correct termstore
            //TaxonomySession session = new TaxonomySession(site);
            //TermStore termStore = session.TermStores[termStoreId];

            try
            {
                if (session.TermStores != null && session.TermStores.Count() > 0)
                {
                    termStore = session.TermStores.FirstOrDefault();
                    logger.Log((string.Format("TermStore {0} presente", termStore.Name.ToString(), LogSeverity.Debug)));
                    return termStore;
                }
                else
                {
                    logger.Log(string.Format("TermStore assente", String.Empty, LogSeverity.Error));
                    return null;
                }
            }
            catch (Exception ex)
            {
                logger.Log(string.Format("TermStore assente", ex.Message, LogSeverity.Error));
                return termStore;
            }
        }
        
        public static bool CheckTermStoreGroup(TermStore termstore, string _group)
        {
            return (termstore.Groups.Any(x => x.Name == _group)) ? true : false;
        }
        
        public static Guid GetIdTermStoreGroup(TermStore termstore, string _group)
        {
            Guid uidStore = new Guid(idStore);
            return (termstore.Groups.Any(x => x.Name == _group)) ? termstore.Groups[_group].Id : uidStore;
        }

        public static bool MakeTermstoreGroup(TermStore termstore, string _group)
        {
            LogHelper logger = LogHelper.Instance;

            bool _result = false;
            try
            {
                Group group = termstore.CreateGroup(_group);
                termstore.CommitAll();
                _result = true;
                logger.Log(string.Format("Gruppo {0} creato nel termstore {1}", _group, termstore.Name.ToString(), LogSeverity.Debug));
            }
            catch (Exception ex)
            {
                _result = false;
                logger.Log(string.Format("Gruppo {0} non creato nel termstore {1}", _group, termstore.Name.ToString() + "  :  " + ex.Message, LogSeverity.Error));
            }
            return _result;
        }

        public static bool CheckTermSet(TermStore termstore, string _group, string _termset)
        {
            return (termstore.GetGroup(GetIdTermStoreGroup(termstore, _group)).TermSets.Any(x => x.Name == _termset)) ? true:false;
        }

        public static Guid GetIdTermSet(TermStore termstore, string _group, string _termset)
        {
            Guid uidStore = new Guid(idStore);
            Group group = termstore.GetGroup(GetIdTermStoreGroup(termstore, _group));
            return (group.TermSets.Any(x => x.Name == _termset)) ? group.TermSets[_termset].Id : uidStore;
        }

        public static bool MakeTermSet(TermStore termstore, string _group, string _termset)
        {
            LogHelper logger = LogHelper.Instance;

            bool _result = false;

            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    Guid termSetId = Guid.NewGuid();
                    Guid uidStore = GetIdTermStoreGroup(termstore, _group);
                    termstore.GetGroup(uidStore).CreateTermSet(_termset, termSetId, termstore.DefaultLanguage);
                    termstore.CommitAll();
                });

                _result = true;
                logger.Log(string.Format("Set Termini {0} creato nel Gruppo {1}", _termset, _group, LogSeverity.Debug));
            }
            catch (Exception ex)
            {
                _result = false;
                logger.Log(ex.Message + " : " + string.Format("Set Termini {0} non creato nel Gruppo {1}", _termset, _group, LogSeverity.Error));
            }
            return _result;            
        }

        public static bool ck_Am(TermStore termstore, string _group, string _termset, string _am)
        {
            TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));
            return (termset.Terms.Any(x => x.Name == _am)) ? true : false;
        }
        public static bool mk_Am(TermStore termstore, string _group, string _termset, string _am)
        {
            LogHelper logger = LogHelper.Instance;
            
            bool _result = false;
            try
            {
                TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));

                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    Term term = termset.CreateTerm(_am, termstore.DefaultLanguage);
                    termset.TermStore.CommitAll();
                });

                _result = true;
                logger.Log(string.Format("Ambito {0} creato nel Set {1}", _am, _termset, LogSeverity.Debug));
            }
            catch (Exception ex)
            {
                _result = false;
                logger.Log(ex.Message + " : " +string.Format("Ambito {0} non creato nel Set {1}", _am, _termset, LogSeverity.Error));
            }
            return _result;
        }
        public static bool sp_Am(TermStore termstore, string _group, string _termset, string _am, string _p, string _v)
        {
            LogHelper logger = LogHelper.Instance;
            
            bool _result = false;

            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    Guid uidStore = new Guid(idStore);
                    TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));
                    Term term = termset.Terms[_am];
                    term.SetCustomProperty(_p, _v);
                    termset.TermStore.CommitAll();
                });

                _result = true;
                logger.Log(string.Format("Proprietà {0} aggiunta nell'ambito {1} ", _p, _am, LogSeverity.Debug));
            }
            catch (Exception ex)
            {
                _result = false;
                logger.Log(ex.Message + " : " + string.Format("Proprietà {0} non aggiunta nell'ambito {1} ", _p, _am, LogSeverity.Error));
            }
            return _result;
        }

        public static bool ck_Pr(TermStore termstore, string _group, string _termset, string _am, string _pr)
        {
            TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));
            return (termset.Terms[_am].Terms.Any(x => x.Name == _pr)) ? true : false;
        }
        public static bool mk_Pr(TermStore termstore, string _group, string _termset, string _am, string _pr)
        {
            LogHelper logger = LogHelper.Instance;
            
            bool _result = false;
            try
            {
                TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));
                Term term = termset.Terms[_am];

                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    term.CreateTerm(_pr, termstore.DefaultLanguage);
                    termset.TermStore.CommitAll();
                });

                _result = true;
                logger.Log(string.Format("Processo {0} creato nell'ambito {1} nel set {2}", _pr, _am, _termset, LogSeverity.Debug));
            }
            catch (Exception ex)
            {
                _result = false;
                logger.Log(ex.Message + " : " + string.Format("Processo {0} non creato nell'ambito {1} nel set {2}", _pr, _am, _termset, LogSeverity.Debug));
            }
            return _result;
        }
        public static bool sp_Pr(TermStore termstore, string _group, string _termset, string _am, string _pr, string _p, string _v)
        {
            LogHelper logger = LogHelper.Instance;
            
            bool _result = false;

            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    Guid uidStore = new Guid(idStore);
                    TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));
                    Term term = termset.Terms[_am].Terms[_pr];
                    term.SetCustomProperty(_p, _v);
                    termset.TermStore.CommitAll();
                });

                _result = true;
                logger.Log(string.Format("Proprietà {0} aggiunta nel processo {1} dell'ambito {2} ", _p, _pr, _am, LogSeverity.Debug));
            }
            catch (Exception ex)
            {
                _result = false;
                logger.Log(ex.Message + " : " + string.Format("Proprietà {0} aggiunta nel processo {1} dell'ambito {2} ", _p, _pr, _am, LogSeverity.Error));
            }
            return _result;
        }

        public static bool ck_At(TermStore termstore, string _group, string _termset, string _am, string _pr, string _at)
        {
            TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));
            return (termset.Terms[_am].Terms[_pr].Terms.Any(x => x.Name == _at)) ? true : false;
        }
        public static bool mk_At(TermStore termstore, string _group, string _termset, string _am, string _pr, string _at)
        {
            LogHelper logger = LogHelper.Instance;
            
            bool _result = false;
            try
            {
                TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));
                Term term = termset.Terms[_am].Terms[_pr];

                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    term.CreateTerm(_at, termstore.DefaultLanguage);
                    termset.TermStore.CommitAll();
                });

                _result = true;
                logger.Log(string.Format("Attivita {0} creata nel processo {1} dell'ambito {2} nel set {3}", _pr, _am, _termset, LogSeverity.Debug));
            }
            catch (Exception ex)
            {
                _result = false;
                logger.Log(ex.Message + " : " + string.Format("Attivita {0} creata nel processo {1} dell'ambito {2} nel set {3}", _at, _pr, _am, _termset, LogSeverity.Debug));
            }
            return _result;
        }
        public static bool sp_At(TermStore termstore, string _group, string _termset, string _am, string _pr, string _at, string _p, string _v)
        {
            LogHelper logger = LogHelper.Instance;

            bool _result = false;

            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    Guid uidStore = new Guid(idStore);
                    TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));
                    Term term = termset.Terms[_am].Terms[_pr].Terms[_at];
                    term.SetCustomProperty(_p, _v);
                    termset.TermStore.CommitAll();
                });

                _result = true;
                logger.Log(string.Format("Proprietà {0} aggiunta nell'attività {1} del processo {2} nell'ambito {3} ", _p, _at, _pr, _am, LogSeverity.Debug));
            }
            catch (Exception ex)
            {
                _result = false;
                logger.Log(ex.Message + " : " + string.Format("Proprietà {0} aggiunta nell'attività {1} del processo {2} nell'ambito {3} ", _p, _at, _pr, _am, LogSeverity.Error));
            }
            return _result;
        }
  
        public static bool ck_NOR(TermStore termstore, string _group, string _termset, string _nor)
        {
            TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));
            return (termset.Terms.Any(x => x.Name == _nor)) ? true : false;
        }
        public static bool mk_NOR(TermStore termstore, string _group, string _termset, string _nor)
        {
            LogHelper logger = LogHelper.Instance;
            
            bool _result = false;
            try
            {
                TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));

                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    Term term = termset.CreateTerm(_nor, termstore.DefaultLanguage);
                    termset.TermStore.CommitAll();
                });

                _result = true;
                logger.Log(string.Format("Normativa {0} creata nel Set {1}", _nor, _termset, LogSeverity.Debug));
            }
            catch (Exception ex)
            {
                _result = false;
                logger.Log(ex.Message + " : " + string.Format("Normativa {0} non creata nel Set {1}", _nor, _termset, LogSeverity.Error));
            }
            return _result;
        }
        public static bool sp_NOR(TermStore termstore, string _group, string _termset, string _nor, string _p, string _v)
        {
            LogHelper logger = LogHelper.Instance;
            
            bool _result = false;

            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    Guid uidStore = new Guid(idStore);
                    TermSet termset = termstore.GetTermSet(GetIdTermSet(termstore, _group, _termset));
                    Term term = termset.Terms[_nor];
                    term.SetCustomProperty(_p, _v);
                    termset.TermStore.CommitAll();
                });

                _result = true;
                logger.Log(string.Format("Proprietà {0} aggiunta nella Normativa {1} ", _p, _nor, LogSeverity.Debug));
            }
            catch (Exception ex)
            {
                _result = false;
                logger.Log(ex.Message + " : " + string.Format("Proprietà {0} non aggiunta nella Normativa {1} ", _p, _nor, LogSeverity.Error));
            }
            return _result;
        }
        
        public static void AssociateMetadata(SPSite site, Guid fieldId, string _MMSGroupName, string _MMSTermsetName)
        {
            string _MMSServiceAppName = "ManagedMetadataServiceApplication";
            if (site.RootWeb.Fields.Contains(fieldId))
            {
                TaxonomySession session = new TaxonomySession(site);
                if (session.TermStores.Count != 0)
                {
                    var termStore = session.TermStores[_MMSServiceAppName];
                    foreach (Group grp in termStore.Groups)
                    {
                        if (grp.Name.ToUpper() == _MMSGroupName.ToUpper())
                        {
                            var group = grp;
                            var termSet = group.TermSets[_MMSTermsetName];
                            TaxonomyField field = site.RootWeb.Fields[fieldId] as TaxonomyField;
                            field.SspId = termSet.TermStore.Id;
                            field.TermSetId = termSet.Id;
                            field.TargetTemplate = string.Empty;
                            field.AnchorId = Guid.Empty;
                            field.Update();
                            break;
                        }
                    }
                }
            }
        }
        public static void DeAssociateMetadata(SPSite currSite, Guid fieldId)
        {
            if (currSite.RootWeb.Fields.Contains(fieldId))
            {
                TaxonomyField field = currSite.RootWeb.Fields[fieldId] as TaxonomyField;

                field.SspId = Guid.Empty;
                field.TermSetId = Guid.Empty;

                field.TargetTemplate = string.Empty;
                field.AnchorId = Guid.Empty;
                field.Update();
            }
        }

        private void SetTaxonomyField(TaxonomySession metadataService, SPListItem item, Guid fieldId, string fieldValue)
        {
            TaxonomyField taxField = item.Fields[fieldId] as TaxonomyField;
            TermStore termStore = metadataService.TermStores[taxField.SspId];
            TermSet termSet = termStore.GetTermSet(taxField.TermSetId);
            SetTaxonomyFieldValue(termSet, taxField, item, fieldValue);
        }

        private void SetTaxonomyField(TaxonomySession metadataService, SPListItem item, Guid fieldId, List<string> fieldValues)
        {
            TaxonomyField taxField = item.Fields[fieldId] as TaxonomyField;
            TermStore termStore = metadataService.TermStores[taxField.SspId];
            TermSet termSet = termStore.GetTermSet(taxField.TermSetId);
            if (taxField.AllowMultipleValues) SetTaxonomyFieldMultiValue(termSet, taxField, item, fieldValues);
            else SetTaxonomyFieldValue(termSet, taxField, item, fieldValues.First());
        }

        private void SetTaxonomyFieldValue(TermSet termSet, TaxonomyField taxField, SPListItem item, string value)
        {
            var terms = termSet.GetTerms(value, true, StringMatchOption.ExactMatch, 1, false);
            if (terms.Count > 0)
            {
                taxField.SetFieldValue(item, terms.First());
            }
        }

        private void SetTaxonomyFieldMultiValue(TermSet termSet, TaxonomyField taxField, SPListItem item, List<string> fieldValues)
        {
            var fieldTerms = new List<Term>();
            foreach (var value in fieldValues)
            {
                var terms = termSet.GetTerms(value, true);
                if (terms.Count > 0)
                {
                    terms.ToList().ForEach(t => fieldTerms.Add(t));
                }
            }
            taxField.SetFieldValue(item, fieldTerms);
        }

        private static void LoadTerms(List<termObject> lstExObjs, Term term, int level)
        {
            if (level > 7)
            {
                return;
            }

            foreach (Term curTerm in term.Terms)
            {
                termObject msExportObject = new termObject();

                msExportObject.AvailableforTagging = (curTerm.IsAvailableForTagging ? "TRUE" : "FALSE");
                msExportObject.TermDescription = (string.IsNullOrEmpty(curTerm.GetDescription())
                                                      ? ""
                                                      : string.Format("\"{0}\"", curTerm.GetDescription()));

                if (level == 2)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = string.Format("\"{0}\"", curTerm.Name);
                }
                else if (level == 3)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = lstExObjs[lstExObjs.Count - 1].Level2Term;
                    msExportObject.Level3Term = string.Format("\"{0}\"", curTerm.Name);
                }
                else if (level == 4)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = lstExObjs[lstExObjs.Count - 1].Level2Term;
                    msExportObject.Level3Term = lstExObjs[lstExObjs.Count - 1].Level3Term;
                    msExportObject.Level4Term = string.Format("\"{0}\"", curTerm.Name);
                }
                else if (level == 5)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = lstExObjs[lstExObjs.Count - 1].Level2Term;
                    msExportObject.Level3Term = lstExObjs[lstExObjs.Count - 1].Level3Term;
                    msExportObject.Level4Term = lstExObjs[lstExObjs.Count - 1].Level4Term;
                    msExportObject.Level5Term = string.Format("\"{0}\"", curTerm.Name);
                }
                else if (level == 6)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = lstExObjs[lstExObjs.Count - 1].Level2Term;
                    msExportObject.Level3Term = lstExObjs[lstExObjs.Count - 1].Level3Term;
                    msExportObject.Level4Term = lstExObjs[lstExObjs.Count - 1].Level4Term;
                    msExportObject.Level5Term = lstExObjs[lstExObjs.Count - 1].Level5Term;
                    msExportObject.Level6Term = string.Format("\"{0}\"", curTerm.Name);
                }
                else if (level == 7)
                {
                    msExportObject.Level1Term = lstExObjs[lstExObjs.Count - 1].Level1Term;
                    msExportObject.Level2Term = lstExObjs[lstExObjs.Count - 1].Level2Term;
                    msExportObject.Level3Term = lstExObjs[lstExObjs.Count - 1].Level3Term;
                    msExportObject.Level4Term = lstExObjs[lstExObjs.Count - 1].Level4Term;
                    msExportObject.Level5Term = lstExObjs[lstExObjs.Count - 1].Level5Term;
                    msExportObject.Level6Term = lstExObjs[lstExObjs.Count - 1].Level6Term;
                    msExportObject.Level7Term = string.Format("\"{0}\"", curTerm.Name);
                }

                lstExObjs.Add(msExportObject);

                if (curTerm.TermsCount > 0)
                {
                    LoadTerms(lstExObjs, curTerm, level + 1);
                }

            }

        }

        public class termObject
        {
            public string AvailableforTagging;
            public string TermDescription;
            public string Level1Term;
            public string Level2Term;
            public string Level3Term;
            public string Level4Term;
            public string Level5Term;
            public string Level6Term;
            public string Level7Term;
        }
    }
}    
    
