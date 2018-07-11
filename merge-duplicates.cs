// Copyright 2018 István Koren
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;

// Implementation of macro editor is preliminary and experimental.
// The Citavi object model is subject to change in future version.

public static class CitaviMacro
{
	public static void Main()
	{

		//if this macro should ALWAYS affect all titles in active project, choose first option
		//if this macro should affect just filtered rows if there is a filter applied and ALL if not, choose second option
		
		//ProjectReferenceCollection references = Program.ActiveProjectShell.Project.References;		
		List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences();
		
		//if we need a ref to the active project
		SwissAcademic.Citavi.Project activeProject = Program.ActiveProjectShell.Project;
		
		
		if (references.Count == 2)
		{
			if (references[0].ReferenceType == references[1].ReferenceType)
			{
				string originalTitle = references[0].Title;
				
				// CreatedOn, check which one is older and then take that CreatedOn and CreatedBy
				if (DateTime.Compare(references[0].CreatedOn, references[1].CreatedOn) > 0)
				{
					// second reference is older
					// CreatedOn is write-protected. We therefore switch the references...
					Reference newer = references[0];
					references[0] = references[1];
					references[1] = newer;
				}
				
				// ModifiedOn is write-protected. It will be updated anyways now.
				
				// Abstract, naive approach...
				if (references[0].Abstract.Text.Trim().Length < references[1].Abstract.Text.Trim().Length)
				{
					references[0].Abstract.Text = references[1].Abstract.Text;
				}
				
				// AccessDate, take newer one
				//TODO: accessdate would need to be parsed
				// right now, we just check if there is one, we take it, otherwise we leave it empty.
				if (references[0].AccessDate.Length < references[1].AccessDate.Length)
				{
					references[0].AccessDate = references[1].AccessDate;
				}
				
				// Additions
				references[0].Additions = MergeOrCombine(references[0].Additions, references[1].Additions);
				
				// CitationKey, check if CitationKeyUpdateType is 0 at one reference if yes, take that one
				if ((references[0].CitationKeyUpdateType == UpdateType.Automatic) && (references[1].CitationKeyUpdateType == UpdateType.Manual))
				{
					references[0].CitationKey = references[1].CitationKey;
					references[0].CitationKeyUpdateType = references[1].CitationKeyUpdateType;
				}
				
				// CoverPath
				if (references[0].CoverPath.LinkedResourceType == LinkedResourceType.Empty)
				{
					references[0].CoverPath = references[1].CoverPath;
				}
				
				// CustomFields (1-9)
				references[0].CustomField1 = MergeOrCombine(references[0].CustomField1, references[1].CustomField1);
				references[0].CustomField2 = MergeOrCombine(references[0].CustomField2, references[1].CustomField2);
				references[0].CustomField3 = MergeOrCombine(references[0].CustomField3, references[1].CustomField3);
				references[0].CustomField4 = MergeOrCombine(references[0].CustomField4, references[1].CustomField4);
				references[0].CustomField5 = MergeOrCombine(references[0].CustomField5, references[1].CustomField5);
				references[0].CustomField6 = MergeOrCombine(references[0].CustomField6, references[1].CustomField6);
				references[0].CustomField7 = MergeOrCombine(references[0].CustomField7, references[1].CustomField7);
				references[0].CustomField8 = MergeOrCombine(references[0].CustomField8, references[1].CustomField8);
				references[0].CustomField9 = MergeOrCombine(references[0].CustomField9, references[1].CustomField9);
				
				// Date (string type
				references[0].Date = MergeOrCombine(references[0].Date, references[1].Date);
				
				// Date2 (string type)
				references[0].Date2 = MergeOrCombine(references[0].Date2, references[1].Date2);
				
				// DOI
				references[0].Doi = MergeOrCombine(references[0].Doi, references[1].Doi);
				
				// Edition
				references[0].Edition = MergeOrCombine(references[0].Edition, references[1].Edition);
				
				// EndPage
				if (references[0].PageRange.ToString() == "")
				{
					references[0].PageRange = references[1].PageRange;
				}
				
				// Evaluation, naive approach...
				if (references[0].Evaluation.Text.Trim().Length < references[1].Evaluation.Text.Trim().Length)
				{
					references[0].Evaluation.Text = references[1].Evaluation.Text;
				}
				
				// HasLabel1 and HasLabel2
				if (references[1].HasLabel1)
				{
					references[0].HasLabel1 = references[1].HasLabel1;
				}
				if (references[1].HasLabel2)
				{
					references[0].HasLabel2 = references[1].HasLabel2;
				}
				
				// ISBN
				references[0].Isbn = MergeOrCombine(references[0].Isbn.ToString(), references[1].Isbn.ToString());
				
				// Language
				references[0].Language = MergeOrCombine(references[0].Language, references[1].Language);
				
				// Notes
				references[0].Notes = MergeOrCombine(references[0].Notes, references[1].Notes);
				
				// Number
				references[0].Number = MergeOrCombine(references[0].Number, references[1].Number);
				
				// NumberOfVolumes
				references[0].NumberOfVolumes = MergeOrCombine(references[0].NumberOfVolumes, references[1].NumberOfVolumes);
				
				// OnlineAddress
				references[0].OnlineAddress = MergeOrCombine(references[0].OnlineAddress, references[1].OnlineAddress);
				
				// OriginalCheckedBy
				references[0].OriginalCheckedBy = MergeOrCombine(references[0].OriginalCheckedBy, references[1].OriginalCheckedBy);
				
				// OriginalPublication
				references[0].OriginalPublication = MergeOrCombine(references[0].OriginalPublication, references[1].OriginalPublication);
				
				// PageCount (text)
				//TODO: apparently it is a calculated field
				// PageCountNumeralSystem (int=0)
				//TODO: apparently it is a calculated field
				// PageRangeNumberingType (int=0)
				//TODO: apparently it is a calculated field
				// PageRangeNumeralSystem (int=0)
				//TODO: apparently it is a calculated field
				
				// ParallelTitle
				references[0].ParallelTitle = MergeOrCombine(references[0].ParallelTitle, references[1].ParallelTitle);

				// PeriodicalID, naive approach...
				if ((references[0].Periodical == null) || (((references[0].Periodical != null) && (references[1].Periodical != null)) && (references[0].Periodical.ToString().Length < references[1].Periodical.ToString().Length)))
				{
					references[0].Periodical = references[1].Periodical;
				}
				
				// PlaceOfPublication
				references[0].PlaceOfPublication = MergeOrCombine(references[0].PlaceOfPublication, references[1].PlaceOfPublication);
				
				// Price
				references[0].Price = MergeOrCombine(references[0].Price, references[1].Price);
				
				// PubMedID
				references[0].PubMedId = MergeOrCombine(references[0].PubMedId, references[1].PubMedId);
				
				// Rating (take average)
				references[0].Rating = (short) Math.Floor((decimal) ((references[0].Rating + references[1].Rating) / 2));
				
				// (!) ReferenceType (not supported)
				
				// SequenceNumber (take the one of first, as second reference will be deleted)
				
				// ShortTitle, check if ShortTitleUpdateType is 0 at one reference if yes, take that one
				if ((references[0].ShortTitleUpdateType == UpdateType.Automatic) && (references[1].ShortTitleUpdateType == UpdateType.Manual))
				{
					references[0].ShortTitle = references[1].ShortTitle;
				}
				else if ((references[0].ShortTitleUpdateType == UpdateType.Manual) && (references[1].ShortTitleUpdateType == UpdateType.Manual))
				{
					references[0].ShortTitle = MergeOrCombine(references[0].ShortTitle, references[1].ShortTitle);
				}
				
				// SourceOfBibliographicInformation
				references[0].SourceOfBibliographicInformation = MergeOrCombine(references[0].SourceOfBibliographicInformation, references[1].SourceOfBibliographicInformation);
				
				// SpecificFields (1-7)
				references[0].SpecificField1 = MergeOrCombine(references[0].SpecificField1, references[1].SpecificField1);
				references[0].SpecificField2 = MergeOrCombine(references[0].SpecificField2, references[1].SpecificField2);
				references[0].SpecificField3 = MergeOrCombine(references[0].SpecificField3, references[1].SpecificField3);
				references[0].SpecificField4 = MergeOrCombine(references[0].SpecificField4, references[1].SpecificField4);
				references[0].SpecificField5 = MergeOrCombine(references[0].SpecificField5, references[1].SpecificField5);
				references[0].SpecificField6 = MergeOrCombine(references[0].SpecificField6, references[1].SpecificField6);
				references[0].SpecificField7 = MergeOrCombine(references[0].SpecificField7, references[1].SpecificField7);
				
				// StartPage
				//TODO: see page range
				
				// StorageMedium
				references[0].StorageMedium = MergeOrCombine(references[0].StorageMedium, references[1].StorageMedium);
				
				// Subtitle
				references[0].Subtitle = MergeOrCombine(references[0].Subtitle, references[1].Subtitle);
				
				// SubtitleTagged
				//TODO: we are not merging SubtitleTagged as that changes the Subtitle as well
				//references[0].SubtitleTagged = MergeOrCombine(references[0].SubtitleTagged, references[1].SubtitleTagged);
				
				// TableOfContents, naive approach...				
				if ((references[0].TableOfContents == null) || (((references[0].TableOfContents != null) && (references[1].TableOfContents != null)) && (references[0].TableOfContents.ToString().Length < references[1].TableOfContents.ToString().Length)))
				{
					references[0].TableOfContents.Text = references[1].TableOfContents.Text;
				}
				
				// TextLinks
				references[0].TextLinks = MergeOrCombine(references[0].TextLinks, references[1].TextLinks);
				
				// Title
				references[0].Title = MergeOrCombine(references[0].Title, references[1].Title);
				
				// TitleTagged
				//TODO: we are not merging TitleTagged as that changes the Title as well
				//references[0].TitleTagged = MergeOrCombine(references[0].TitleTagged, references[1].TitleTagged);
				
				// TitleInOtherLanguages
				references[0].TitleInOtherLanguages = MergeOrCombine(references[0].TitleInOtherLanguages, references[1].TitleInOtherLanguages);
				
				// TitleSupplement
				references[0].TitleSupplement = MergeOrCombine(references[0].TitleSupplement, references[1].TitleSupplement);
				
				// TitleSupplementTagged
				//TODO: we are not merging TitleSupplementTagged as that changes the TitleSupplement as well
				//references[0].TitleSupplementTagged = MergeOrCombine(references[0].TitleSupplementTagged, references[1].TitleSupplementTagged);
				
				// TranslatedTitle
				references[0].TranslatedTitle = MergeOrCombine(references[0].TranslatedTitle, references[1].TranslatedTitle);
				
				// UniformTitle
				references[0].UniformTitle = MergeOrCombine(references[0].UniformTitle, references[1].UniformTitle);
				
				// Volume
				references[0].Volume = MergeOrCombine(references[0].Volume, references[1].Volume);
				
				// Year
				references[0].Year = MergeOrCombine(references[0].Year, references[1].Year);
				
				// ReservedData
				//TODO: apparently cannot be set
				
				// RecordVersion (?)
				//TODO: apparently cannot be set
				
				
				// FOREIGN KEY fields
				
				// Locations
				references[0].Locations.AddRange(references[1].Locations);
				
				// Groups
				references[0].Groups.AddRange(references[1].Groups);
				
				// Quotations
				references[0].Quotations.AddRange(references[1].Quotations);
				
				// ReferenceAuthors
				//references[0].Authors.AddRange(references[1].Authors);
				foreach (Person author in references[1].Authors)
				{
					if (!references[0].Authors.Contains(author))
					{
						references[0].Authors.Add(author);
					}
				}
				
				// ReferenceCategory
				references[0].Categories.AddRange(references[1].Categories);
				
				// ReferenceCollaborator
				references[0].Collaborators.AddRange(references[1].Collaborators);
				
				// ReferenceEditor
				references[0].Editors.AddRange(references[1].Editors);
				
				// ReferenceKeyword
				references[0].Keywords.AddRange(references[1].Keywords);
				
				// ReferenceOrganization
				references[0].Organizations.AddRange(references[1].Organizations);
				
				// ReferenceOthersInvolved
				references[0].OthersInvolved.AddRange(references[1].OthersInvolved);
				
				// ReferencePublisher
				references[0].Publishers.AddRange(references[1].Publishers);
				
				// ReferenceReference
				// adding ChildReferences does not work
				//references[0].ChildReferences.AddRange(references[1].ChildReferences);
				Reference[] childReferences = references[1].ChildReferences.ToArray();
				foreach (Reference child in childReferences)
				{
					child.ParentReference = references[0];
				}
				
				// SeriesTitle, naive approach
				if ((references[0].SeriesTitle == null) || (((references[0].SeriesTitle != null) && (references[1].SeriesTitle != null)) && (references[0].SeriesTitle.ToString().Length < references[1].SeriesTitle.ToString().Length)))
				{
					references[0].SeriesTitle = references[1].SeriesTitle;
				}
				
				
				// change crossreferences
				foreach (EntityLink entityLink in references[1].EntityLinks)
         		{
					if (entityLink.Source == references[1])
					{
						entityLink.Source = references[0];
					}
					else if (entityLink.Target == references[1])
					{
						entityLink.Target = references[0];
					}
				}
				
				
				// write Note that the reference has been merged
				if (references[0].Notes.Trim().Length > 0)
				{
					references[0].Notes += " |";
				}
				references[0].Notes += " This reference has been merged with a duplicate by CitaviBot.";
				
				// DONE! remove second reference
				activeProject.References.Remove(references[1]);

			}
			else
			{
				MessageBox.Show("Currently this script only supports merging two references of the same type. Please convert and try again.");
			}
		}
		else
		{
			MessageBox.Show("Currently this script only supports merging two references. Please select two and try again.");
		}
	
		

		foreach (Reference reference in references)
		{
			ReferenceType referenceType = reference.ReferenceType;
			
			//if you need to operate on references of a certain reference type only:
			//if (reference.ReferenceType == ReferenceType.Book) ...
			
		}
		
	}
	
	private static string MergeOrCombine(string first, string second) {
		first = first.Trim();
		second = second.Trim();
		
		// do not compare ignore case, otherwise we might lose capitalization information; in that case we rely on manual edits after the merge
		if (String.Compare(first, second, false) == 0)
		{
			// easy case, they are the same!
			return first;
		}
		else if (first.Length == 0)
		{
			return second;
		}
		else if (second.Length == 0)
		{
			return first;
		}
		else
		{
			return first + " // " + second;
		}
	}
	
	private static void MergeReferencePersonCollections(ReferencePersonCollection first, ReferencePersonCollection second)
	{
		List<Person> personsList = new List<Person>();
		personsList.AddRange(first);
		Person[] personsArray = second.ToArray();
		foreach (Person person in personsArray)
		{
			personsList.AddIfNotExists(person);
		}
		first.ReplaceBy(personsList);
	}
	
	private static DateTime GetOlderDateTime(DateTime first, DateTime second)
	{
		if (DateTime.Compare(first, second) < 0)
		{
			return first;
		}
		else
		{
			return second;
		}
	}
	
}