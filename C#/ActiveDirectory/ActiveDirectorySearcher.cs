/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
namespace Data.ActiveDirectory
{
    using System;
    using System.Text;
    using System.Collections.Generic;
    using System.DirectoryServices;

    /// <summary>
    /// Base class for searching Active Directory
    /// </summary>
    public class ActiveDirectorySearcher
    {
        /// <summary>
        /// Default AD global catalog for searching
        /// </summary>
        private const string DefaultGlobalCatalog = "GC://corp.microsoft.com";

        /// <summary>
        /// Active Directory properties used in this project
        /// </summary>
        protected static readonly string[] AdProperties = new[]
            {
                "mailnickname", "displayname", "mail", "givenname", "sn", "telephoneNumber", "title", "physicalDeliveryOfficeName",
                "department", "distinguishedname", "manager", "directreports", "thumbnailphoto"
            };

        /// <summary>
        /// DirectorySearcher for the class
        /// </summary>
        private readonly DirectorySearcher ds;

        /// <summary>
        /// Initializes a new instance of the <see cref="ActiveDirectorySearcher"/> class.
        /// Use derived classes for public interface.
        /// </summary>
        protected ActiveDirectorySearcher()
        {
            this.ds = new DirectorySearcher(new DirectoryEntry(DefaultGlobalCatalog));
        }

        /// <summary>
        /// Gets the directory searcher.
        /// </summary>
        /// <value>The ds.</value>
        protected DirectorySearcher Ds
        {
            get
            {
                return this.ds;
            }
        }

        /// <summary>
        /// Gets domain from Distinguished Name.
        /// </summary>
        /// <param name="distinguishedName">The Distinguished Name.</param>
        /// <returns>Domain for Distinguished Name</returns>
        protected static string DomainFromDistinguishedName(string distinguishedName)
        {
            string domain = null;
            int index = distinguishedName.IndexOf("DC=");

            if (index > 0)
            {
                domain = distinguishedName.Substring(index + 3);       // to skip the "DC="

                index = domain.IndexOf(",");
                if (index > 0)
                {
                    domain = domain.Substring(0, index);
                }
            }

            return domain;
        }

        /// <summary>
        /// Escapes the given LDAP string by adding 
        /// a backslash in front of (, ), and \.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns>Escaped string</returns>
        protected static string EscapeLdap(string text)
        {
            var output = new StringBuilder();

            foreach (char ch in text)
            {
                switch (ch)
                {
                    case '(':
                        output.Append("\\28");
                        break;
                    case ')':
                        output.Append("\\29");
                        break;
                    case '\\':
                        output.Append("\\5c");
                        break;
                    default:
                        output.Append(ch);
                        break;
                }
            }

            return output.ToString();
        }

        /// <summary>
        /// Gets the value of the given property from the given search result.
        /// </summary>
        /// <param name="result">The result.</param>
        /// <param name="property">The property to look for in the result.</param>
        /// <returns>If found, property value cast to a string; otherwise empty string</returns>
        protected static string GetResultProperty(SearchResult result, string property)
        {
            // Use the property argument to index into Properties, which is 
            // a ResultPropertyCollection. Such indexing is supported in C#.
            // Actually using the Item property of the ResultPropertyCollection.
            if (result != null && result.Properties[property] != null && 
                result.Properties[property].Count > 0)
            {
                return (string)result.Properties[property][0];
            }

            // Return empty (instead of NULL) because ASPX Literal controls can't deal with null
            return string.Empty;
        }

        /// <summary>
        /// Create Person object from search result.
        /// </summary>
        /// <param name="result">The result.</param>
        /// <returns>Person object or null</returns>
        protected Person PersonFromSearchResult(SearchResult result)
        {
            if (result == null)
            {
                return null;
            }

            var person = new Person
            {
                DisplayName = GetResultProperty(result, "displayname"),
                EmailAddress = GetResultProperty(result, "mail"),
                FirstName = GetResultProperty(result, "givenname"),
                LastName = GetResultProperty(result, "sn"),
                Telephone = GetResultProperty(result, "telephoneNumber"),
                Title = GetResultProperty(result, "title"),
                Alias = GetResultProperty(result, "mailnickname"),
                Office = GetResultProperty(result, "physicalDeliveryOfficeName"),
                Department = GetResultProperty(result, "department"),
                DistinguishedName = GetResultProperty(result, "distinguishedname"),
                Manager = GetResultProperty(result, "manager"),
                Directs = null,
                ThumbnailPhoto = null
            };

            person.ThumbnailPhoto = ThumbnailFromSearchResult(result);
            if (person.ThumbnailPhoto != null)
            {
                // Encoded thumbnail to be used in data:image/jpg;base64.
                person.EncodedThumbnail = Convert.ToBase64String(person.ThumbnailPhoto);
            }

            person.DirectsCount = (result.Properties["directreports"] != null) ? result.Properties["directreports"].Count : 0;
            if (person.DirectsCount > 0)
            {
                person.Directs = new List<string>(person.DirectsCount);
                for(int i = 0; i < person.DirectsCount; i++)
                {
                    person.Directs.Add(result.Properties["directreports"][i].ToString());
                }
            }

            return person;
        }

        /// <summary>
        /// Helper function to get thumbnail from search result.
        /// </summary>
        /// <param name="result">The results.</param>
        /// <returns>Byte array representing the thumbnail.</returns>
        protected byte[] ThumbnailFromSearchResult(SearchResult result)
        {
            byte[] thumbnail = null;

            if (result.Properties["thumbnailphoto"] != null && result.Properties["thumbnailphoto"].Count > 0)
            {
                thumbnail = (byte[])result.Properties["thumbnailphoto"][0];

                if (thumbnail != null)
                {
                    // approximate jpg check
                    if (thumbnail.Length > 16
                        && thumbnail[6] == 'J' && thumbnail[7] == 'F'
                        && thumbnail[8] == 'I' && thumbnail[9] == 'F')
                    {
                        // already jpg
                    }
                    else
                    {
                        try
                        {
                            var temp = Convert.FromBase64String(Encoding.ASCII.GetString(thumbnail));
                            thumbnail = temp;
                        }
                        catch (FormatException)
                        {
                            // skip on exception
                        }
                    }
                }
            }

            if (thumbnail == null)
            {
                using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
                {
                    Resource.anonymous.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                    thumbnail = ms.ToArray();
                }
            }

            return thumbnail;
        }
    }
}
// *********************************************************
//
// Outlook-Add-in-Display-Info-From-AD, https://github.com/OfficeDev/Outlook-Add-in-Display-Info-From-AD
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************