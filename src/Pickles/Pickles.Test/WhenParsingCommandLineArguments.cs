﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NUnit.Framework;
using Pickles.Extensions;
using Should;

namespace Pickles.Test
{
    [TestFixture]
    public class WhenParsingCommandLineArguments
    {
        private static readonly string expectedHelpString = @"  -f, --feature-directory=VALUE
                             directory to start scanning recursively for 
                               features
  -o, --output-directory=VALUE
                             directory where output files will be placed
      --trfmt, --test-results-format=VALUE
                             the format of the linked test results 
                               (nunit|xunit)
      --lr, --link-results-file=VALUE
                             the path to the linked test results file
      --sn, --system-under-test-name=VALUE
                             a file containing the results of testing the 
                               features
      --sv, --system-under-test-version=VALUE
                             the name of the system under test
  -l, --language=VALUE       the language of the feature files
      --df, --documentation-format=VALUE
                             the format of the output documentation
  -v, --version              
  -h, -?, --help";

        private static readonly string expectedVersionString = 
                string.Format(@"Pickles version {0}", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version);

        [Test]
        public void Then_can_parse_short_form_arguments_successfully()
        {
            var args = new string[] { @"-f=c:\features", @"-o=c:\features-output" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            Assert.AreEqual(true, shouldContinue);
            Assert.AreEqual(@"c:\features", configuration.FeatureFolder.FullName);
            Assert.AreEqual(@"c:\features-output", configuration.OutputFolder.FullName);
            Assert.AreEqual(false, configuration.HasTestResults);
            Assert.AreEqual(null, configuration.TestResultsFile);
        }

        [Test]
        public void Then_can_parse_long_form_arguments_successfully()
        {
            var args = new string[] { @"--feature-directory=c:\features", @"--output-directory=c:\features-output" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            Assert.AreEqual(true, shouldContinue);
            Assert.AreEqual(@"c:\features", configuration.FeatureFolder.FullName);
            Assert.AreEqual(@"c:\features-output", configuration.OutputFolder.FullName);
            Assert.AreEqual(false, configuration.HasTestResults);
            Assert.AreEqual(null, configuration.TestResultsFile);
        }

        [Test]
        public void Then_can_parse_results_format_nunit_with_short_form_successfully()
        {
            var args = new string[] { @"-trfmt=nunit" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            Assert.AreEqual(true, shouldContinue);
            Assert.AreEqual(TestResultsFormat.NUnit, configuration.TestResultsFormat);
        }

        [Test]
        public void Then_can_parse_results_format_xunit_with_short_form_successfully()
        {
            var args = new string[] { @"-trfmt=xunit" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            Assert.AreEqual(true, shouldContinue);
            Assert.AreEqual(TestResultsFormat.xUnit, configuration.TestResultsFormat);
        }

        [Test]
        public void Then_can_parse_results_format_mstest_with_short_form_successfully()
        {
            var args = new string[] { @"-trfmt=mstest" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            Assert.AreEqual(true, shouldContinue);
            Assert.AreEqual(TestResultsFormat.MsTest, configuration.TestResultsFormat);
        }

        [Test]
        public void Then_can_parse_results_format_nunit_with_long_form_successfully()
        {
            var args = new string[] { @"-test-results-format=nunit" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            Assert.AreEqual(true, shouldContinue);
            Assert.AreEqual(TestResultsFormat.NUnit, configuration.TestResultsFormat);
        }

        [Test]
        public void Then_can_parse_results_format_xunit_with_long_form_successfully()
        {
            var args = new string[] { @"-test-results-format=xunit" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            Assert.AreEqual(true, shouldContinue);
            Assert.AreEqual(TestResultsFormat.xUnit, configuration.TestResultsFormat);
        }

        [Test]
        public void Then_can_parse_results_format_mstest_with_long_form_successfully()
        {
            var args = new string[] { @"-test-results-format=mstest" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            Assert.AreEqual(true, shouldContinue);
            Assert.AreEqual(TestResultsFormat.MsTest, configuration.TestResultsFormat);
        }

        [Test]
        public void Then_can_parse_results_file_with_short_form_successfully()
        {
            var args = new string[] { @"-lr=c:\results.xml" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            Assert.AreEqual(true, shouldContinue);
            Assert.AreEqual(Path.GetFullPath(Directory.GetCurrentDirectory()), configuration.FeatureFolder.FullName);
            Assert.AreEqual(Path.GetFullPath(Environment.GetEnvironmentVariable("TEMP")), configuration.OutputFolder.FullName);
            Assert.AreEqual(true, configuration.HasTestResults);
            Assert.AreEqual(@"c:\results.xml", configuration.TestResultsFile.FullName);
        }

        [Test]
        public void Then_can_parse_results_file_with_long_form_successfully()
        {
            var args = new string[] { @"-link-results-file=c:\results.xml" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            Assert.AreEqual(true, shouldContinue);
            Assert.AreEqual(Path.GetFullPath(Directory.GetCurrentDirectory()), configuration.FeatureFolder.FullName);
            Assert.AreEqual(Path.GetFullPath(Environment.GetEnvironmentVariable("TEMP")), configuration.OutputFolder.FullName);
            Assert.AreEqual(true, configuration.HasTestResults);
            Assert.AreEqual(@"c:\results.xml", configuration.TestResultsFile.FullName);
        }

        [Test]
        public void Then_can_parse_excel_documentation_format_with_short_form_successfully()
        {
            var args = new string[] { @"-df=excel" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            shouldContinue.ShouldBeTrue();
            configuration.DocumentationFormat.ShouldEqual(DocumentationFormat.Excel);
        }

        [Test]
        public void Then_can_parse_excel_documentation_format_with_long_form_successfully()
        {
            var args = new string[] { @"-documentation-format=excel" };

            var configuration = new Configuration();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, TextWriter.Null);

            shouldContinue.ShouldBeTrue();
            configuration.DocumentationFormat.ShouldEqual(DocumentationFormat.Excel);
        }

        [Test]
        public void Then_can_parse_help_request_with_question_mark_successfully()
        {
            var args = new string[] { @"-?" };

            var configuration = new Configuration();
            var writer = new StringWriter();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, writer);

            StringAssert.Contains(expectedHelpString.ComparisonNormalize(), writer.GetStringBuilder().ToString().ComparisonNormalize());
            Assert.AreEqual(false, shouldContinue);
            Assert.AreEqual(Path.GetFullPath(Directory.GetCurrentDirectory()), configuration.FeatureFolder.FullName);
            Assert.AreEqual(Path.GetFullPath(Environment.GetEnvironmentVariable("TEMP")), configuration.OutputFolder.FullName);
            Assert.AreEqual(false, configuration.HasTestResults);
            Assert.AreEqual(null, configuration.TestResultsFile);
        }

        [Test]
        public void Then_can_parse_help_request_with_short_form_successfully()
        {
            var args = new string[] { @"-h" };

            var configuration = new Configuration();
            var writer = new StringWriter();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, writer);

            StringAssert.Contains(expectedHelpString.ComparisonNormalize(), writer.GetStringBuilder().ToString().ComparisonNormalize());
            Assert.AreEqual(false, shouldContinue);
            Assert.AreEqual(Path.GetFullPath(Directory.GetCurrentDirectory()), configuration.FeatureFolder.FullName);
            Assert.AreEqual(Path.GetFullPath(Environment.GetEnvironmentVariable("TEMP")), configuration.OutputFolder.FullName);
            Assert.AreEqual(false, configuration.HasTestResults);
            Assert.AreEqual(null, configuration.TestResultsFile);
        }

        [Test]
        public void Then_can_parse_help_request_with_long_form_successfully()
        {
            var args = new[] { @"--help" };

            var configuration = new Configuration();
            var writer = new StringWriter();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, writer);


            var expectedVersionAndHelp = expectedVersionString + Environment.NewLine + expectedHelpString;

            StringAssert.Contains(expectedHelpString.ComparisonNormalize(), writer.GetStringBuilder().ToString().ComparisonNormalize());
            Assert.AreEqual(false, shouldContinue);
            Assert.AreEqual(Path.GetFullPath(Directory.GetCurrentDirectory()), configuration.FeatureFolder.FullName);
            Assert.AreEqual(Path.GetFullPath(Environment.GetEnvironmentVariable("TEMP")), configuration.OutputFolder.FullName);
            Assert.AreEqual(false, configuration.HasTestResults);
            Assert.AreEqual(null, configuration.TestResultsFile);
        }

        [Test]
        public void Then_can_parse_version_request_short_form_successfully()
        {
            var args = new string[] { @"-v" };

            var configuration = new Configuration();
            var writer = new StringWriter();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, writer);

            StringAssert.IsMatch(expectedVersionString.ComparisonNormalize(), writer.GetStringBuilder().ToString().ComparisonNormalize());
            Assert.AreEqual(false, shouldContinue);
            Assert.AreEqual(Path.GetFullPath(Directory.GetCurrentDirectory()), configuration.FeatureFolder.FullName);
            Assert.AreEqual(Path.GetFullPath(Environment.GetEnvironmentVariable("TEMP")), configuration.OutputFolder.FullName);
            Assert.AreEqual(false, configuration.HasTestResults);
            Assert.AreEqual(null, configuration.TestResultsFile);
        }

        [Test]
        public void Then_can_parse_version_request_long_form_successfully()
        {
            var args = new string[] { @"--version" };

            var configuration = new Configuration();
            var writer = new StringWriter();
            var commandLineArgumentParser = new CommandLineArgumentParser();
            bool shouldContinue = commandLineArgumentParser.Parse(args, configuration, writer);

            StringAssert.IsMatch(expectedVersionString.ComparisonNormalize(), writer.GetStringBuilder().ToString().ComparisonNormalize());
            Assert.AreEqual(false, shouldContinue);
            Assert.AreEqual(Path.GetFullPath(Directory.GetCurrentDirectory()), configuration.FeatureFolder.FullName);
            Assert.AreEqual(Path.GetFullPath(Environment.GetEnvironmentVariable("TEMP")), configuration.OutputFolder.FullName);
            Assert.AreEqual(false, configuration.HasTestResults);
            Assert.AreEqual(null, configuration.TestResultsFile);
        }
    }
}
