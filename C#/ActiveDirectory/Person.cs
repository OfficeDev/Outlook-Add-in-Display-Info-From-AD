/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
namespace Data.ActiveDirectory
{
    using System;
    using System.IO;

    using System.Xml;
    using System.Xml.Serialization;

    using System.Collections.Generic;

    using System.Diagnostics;

    /// <summary>
    /// Class reperesents a person.
    /// </summary>
    [Serializable]
    public class Person
    {
        public string Alias;

        public string FirstName;
        public string LastName;
        public string DisplayName;

        public string EmailAddress;
        public string Title;
        public string Office;
        public string Telephone;

        public int DirectsCount;
        public string Department;

        [XmlIgnore]
        public string DistinguishedName;

        [XmlIgnore]
        public string Manager;

        [XmlIgnore]
        public List<string> Directs;

        [XmlIgnore]
        public byte[] ThumbnailPhoto;
        public string EncodedThumbnail;

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