﻿using System;
using Microsoft.Build.Framework;
using ExcelDna.AddIn.Tasks.Logging;
using ExcelDna.AddIn.Tasks.Utils;
using System.IO;

namespace ExcelDna.AddIn.Tasks
{
    public class PackExcelAddIn : AbstractTask
    {
        private readonly IBuildLogger _log;
        private readonly IExcelDnaFileSystem _fileSystem;

        public PackExcelAddIn()
        {
            _log = new BuildLogger(this, "PackExcelAddIn");
            _fileSystem = new ExcelDnaPhysicalFileSystem();
        }

        internal PackExcelAddIn(IBuildLogger log, IExcelDnaFileSystem fileSystem)
        {
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
        }

        public override bool Execute()
        {
            try
            {
                _log.Debug("Running PackExcelAddIn Task");

                int result = ExcelDna.PackedResources.ExcelDnaPack.Pack(OutputDnaFileName, OutputPackedXllFileName, CompressResources, RunMultithreaded, true, null, null);
                if (result != 0)
                    throw new ApplicationException($"Pack failed with exit code {result}.");

                if (SignTool != null && SignOptions != null)
                    Utils.SignTool.Sign(SignTool, SignOptions, OutputPackedXllFileName, _log);

                return true;
            }
            catch (Exception ex)
            {
                _log.Error(ex, ex.Message);
                _log.Error(ex, ex.ToString());
                return false;
            }
        }

        public static string GetOutputPackedXllFileName(string outputXllFileName, string packedFileName, string packedFileSuffix, string publishPath)
        {
            string outputPackedXllFileName = outputXllFileName;
            if (string.IsNullOrWhiteSpace(packedFileName) && !string.IsNullOrWhiteSpace(packedFileSuffix))
            {
                if (IsNone(packedFileSuffix))
                    packedFileSuffix = "";
                packedFileName = Path.GetFileNameWithoutExtension(outputXllFileName) + packedFileSuffix;
            }
            if (!string.IsNullOrWhiteSpace(packedFileName))
            {
                outputPackedXllFileName = Path.Combine(publishPath, packedFileName + ".xll");
            }
            return outputPackedXllFileName;
        }

        public static bool NoPublishPath(string publishPath)
        {
            return IsNone(publishPath);
        }

        public static string GetPublishDirectory(string outDirectory, string publishPath)
        {
            if (NoPublishPath(publishPath))
                return outDirectory;

            return Path.Combine(outDirectory, publishPath ?? "publish");
        }

        private static bool IsNone(string s)
        {
            return string.Equals(s, "%none%", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// The path to the primary .dna file for the ExcelDna add-in
        /// </summary>
        [Required]
        public string OutputDnaFileName { get; set; }

        /// <summary>
        /// Output path
        /// </summary>
        [Required]
        public string OutputPackedXllFileName { get; set; }

        /// <summary>
        /// Compress (LZMA) of resources
        /// </summary>
        [Required]
        public bool CompressResources { get; set; }

        /// <summary>
        /// Use multi threading
        /// </summary>
        [Required]
        public bool RunMultithreaded { get; set; }

        /// <summary>
        /// Path to signtool.exe
        /// </summary>
        public string SignTool { get; set; }

        /// <summary>
        /// Options for signtool.exe
        /// </summary>
        public string SignOptions { get; set; }
    }
}
