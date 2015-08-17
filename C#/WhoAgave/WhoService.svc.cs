/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
using System.IO;
namespace Service.Who
{
    public class WhoService : IWhoService
    {
        /// <summary>
        /// Active Directory source object
        /// </summary>
        private static Data.ActiveDirectory.ActiveDirectorySource ads = new Data.ActiveDirectory.ActiveDirectorySource();

        /// <summary>
        /// Wrapper method to return a person from the given SMTP address.
        /// </summary>
        /// <param name="emailAddress">An SMTP email address.</param>
        /// <returns>PersonContext object.</returns>
        public Data.ActiveDirectory.PersonContext FindPerson(string emailAddress)
        {
            if (string.IsNullOrEmpty(emailAddress)) return null;

            try
            {
                Data.ActiveDirectory.PersonContext person = ads.FindPersonContextBySMTPAddress(emailAddress);

                return person;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Wrapper method to return a person thumbnail.
        /// </summary>
        /// <param name="emailAddress">An SMTP email address.</param>
        /// <returns></returns>
        public System.IO.Stream GetImage(string emailAddress)
        {
            if (string.IsNullOrEmpty(emailAddress)) return null;

            try
            {
                Data.ActiveDirectory.PersonContext person = ads.FindPersonContextBySMTPAddress(emailAddress);

                if (person.Person.ThumbnailPhoto != null)
                {
                    Stream ms = new MemoryStream(person.Person.ThumbnailPhoto);
                    return ms;
                }
            }
            catch
            {
                return null;
            }

            return null;
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