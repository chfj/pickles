﻿#region License

/*
    Copyright [2011] [Jeffrey Cameron]

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using Pickles.DirectoryCrawler;
using Pickles.Extensions;
using Pickles.Parser;

namespace Pickles.DocumentationBuilders.Word
{
    public class WordFeatureFormatter
    {
        private readonly WordScenarioFormatter wordScenarioFormatter;
        private readonly WordScenarioOutlineFormatter wordScenarioOutlineFormatter;
        private readonly WordStyleApplicator wordStyleApplicator;

        public WordFeatureFormatter(WordScenarioFormatter wordScenarioFormatter, WordScenarioOutlineFormatter wordScenarioOutlineFormatter, WordStyleApplicator wordStyleApplicator)
        {
            this.wordScenarioFormatter = wordScenarioFormatter;
            this.wordScenarioOutlineFormatter = wordScenarioOutlineFormatter;
            this.wordStyleApplicator = wordStyleApplicator;
        }

        public void Format(Body body, FeatureDirectoryTreeNode featureDirectoryTreeNode)
        {
            var feature = featureDirectoryTreeNode.Feature;

            body.GenerateParagraph(feature.Name, "Heading2");
            body.GenerateParagraph(feature.Description, "Normal");

            foreach (var featureElement in feature.FeatureElements)
            {
                var scenario = featureElement as Scenario;
                if (scenario != null)
                {
                    this.wordScenarioFormatter.Format(body, scenario);
                }

                var scenarioOutline = featureElement as ScenarioOutline;
                if (scenarioOutline != null)
                {
                    this.wordScenarioOutlineFormatter.Format(body, scenarioOutline);
                }
            }
        }
    }
}