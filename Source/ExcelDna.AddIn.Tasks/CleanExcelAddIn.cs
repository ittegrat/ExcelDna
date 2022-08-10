﻿using System;
using System.Collections;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using ExcelDna.AddIn.Tasks.Logging;
using ExcelDna.AddIn.Tasks.Utils;

namespace ExcelDna.AddIn.Tasks
{
    public class CleanExcelAddIn : AbstractTask
    {
        private readonly IBuildLogger _log;
        private readonly IExcelDnaFileSystem _fileSystem;
        private BuildTaskCommon _common;

        public CleanExcelAddIn()
        {
            _log = new BuildLogger(this, "ExcelDnaClean");
            _fileSystem = new ExcelDnaPhysicalFileSystem();
        }

        internal CleanExcelAddIn(IBuildLogger log, IExcelDnaFileSystem fileSystem)
        {
            _log = log ?? throw new ArgumentNullException(nameof(log));
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
        }

        public override bool Execute()
        {
            try
            {
                _log.Debug("Running CleanExcelAddIn MSBuild Task");

                LogDiagnostics();

                FilesInProject = FilesInProject ?? new ITaskItem[0];
                _log.Debug("Number of files in project: " + FilesInProject.Length);

                _common = new BuildTaskCommon(FilesInProject, OutDirectory, FileSuffix32Bit, FileSuffix64Bit, ProjectName, AddInFileName);

                var existingBuiltFiles = _common.GetBuildItemsForDnaFiles();
                List<ITaskItem> packedFilesToDelete = GetPackedFilesToDelete(existingBuiltFiles);

                // Get the packed name versions : Refactor this + build items
                DeleteAddInFiles(existingBuiltFiles);

                if (PackExcelAddIn.NoPublishPath(PublishPath))
                {
                    DeletePackedAddInFiles(packedFilesToDelete);
                    if (UnpackIsEnabled)
                        DeleteUnpackedAddInFiles();
                }
                else
                {
                    DeletePublishedFiles();
                }

                return true;
            }
            catch (Exception ex)
            {
                _log.Error(ex, ex.Message);
                _log.Error(ex, ex.ToString());
                return false;
            }
        }

        private void LogDiagnostics()
        {
            _log.Debug("----Arguments----");
            _log.Debug("FilesInProject: " + (FilesInProject ?? new ITaskItem[0]).Length);

            if (FilesInProject != null)
            {
                foreach (var f in FilesInProject)
                {
                    _log.Debug($"  {f.ItemSpec}");
                }
            }

            _log.Debug("OutDirectory: " + OutDirectory);
            _log.Debug("Xll32FilePath: " + Xll32FilePath);
            _log.Debug("Xll64FilePath: " + Xll64FilePath);
            _log.Debug("FileSuffix32Bit: " + FileSuffix32Bit);
            _log.Debug("FileSuffix64Bit: " + FileSuffix64Bit);
            _log.Debug("-----------------");
        }

        private List<ITaskItem> GetPackedFilesToDelete(BuildItemSpec[] existingBuiltFiles)
        {
            var packedFilesToDelete = new List<ITaskItem>();

            foreach (var item in existingBuiltFiles)
            {
                packedFilesToDelete.Add(GetPackedFileNames(item.OutputDnaFileNameAs32Bit, item.OutputXllFileNameAs32Bit, item.OutputConfigFileNameAs32Bit, Packed32BitXllName));
                packedFilesToDelete.Add(GetPackedFileNames(item.OutputDnaFileNameAs64Bit, item.OutputXllFileNameAs64Bit, item.OutputConfigFileNameAs64Bit, Packed64BitXllName));
            }

            return packedFilesToDelete;
        }

        private TaskItem GetPackedFileNames(string outputDnaFileName, string outputXllFileName, string outputXllConfigFileName, string packedFileName)
        {
            var outputPackedDnaFileName = !string.IsNullOrWhiteSpace(PackedFileSuffix)
            ? Path.Combine(Path.GetDirectoryName(outputDnaFileName) ?? string.Empty,
                Path.GetFileNameWithoutExtension(outputDnaFileName) + PackedFileSuffix + ".dna")
            : outputDnaFileName;

            var outputPackedXllFileName = PackExcelAddIn.GetOutputPackedXllFileName(outputXllFileName, packedFileName, PackedFileSuffix, PackExcelAddIn.GetPublishDirectory(OutDirectory, PublishPath));

            var outputPackedXllConfigFileName = outputPackedXllFileName + ".config";

            var metadata = new Hashtable
            {
                {"OutputPackedDnaFileName", outputPackedDnaFileName},
                {"OutputPackedXllFileName", outputPackedXllFileName},
                {"OutputPackedXllConfigFileName", outputPackedXllConfigFileName },
            };

            return new TaskItem(outputDnaFileName, metadata);
        }

        private void DeleteAddInFiles(BuildItemSpec[] buildItemsForDnaFiles)
        {
            foreach (var item in buildItemsForDnaFiles)
            {
                DeleteFileIfExists(item.OutputDnaFileNameAs32Bit);
                DeleteFileIfExists(item.OutputDnaFileNameAs64Bit);

                DeleteFileIfExists(item.OutputXllFileNameAs32Bit);
                DeleteFileIfExists(item.OutputXllFileNameAs64Bit);

                DeleteFileIfExists(item.OutputConfigFileNameAs32Bit);
                DeleteFileIfExists(item.OutputConfigFileNameAs64Bit);
            }
        }

        private void DeletePublishedFiles()
        {
            string publishDir = PackExcelAddIn.GetPublishDirectory(OutDirectory, PublishPath);
            if (Directory.Exists(publishDir))
                Array.ForEach(Directory.GetFiles(publishDir), _fileSystem.DeleteFile);
        }

        private void DeletePackedAddInFiles(List<ITaskItem> filesToDelete)
        {
            filesToDelete.ToList().ForEach(f =>
            {
                DeleteFileIfExists(f.GetMetadata("OutputPackedDnaFileName"));
                DeleteFileIfExists(f.GetMetadata("OutputPackedXllFileName"));
                DeleteFileIfExists(f.GetMetadata("OutputPackedXllConfigFileName"));
            });
        }

        private void DeleteUnpackedAddInFiles()
        {
            string[] assemblies = { "ExcelDna.ManagedHost", "ExcelDna.Integration", "ExcelDna.Loader" };
            foreach (var i in assemblies)
            {
                DeleteFileIfExists(Path.Combine(OutDirectory, i + ".dll"));
                DeleteFileIfExists(Path.Combine(OutDirectory, i + ".pdb"));
            }
        }

        private void DeleteFileIfExists(string path)
        {
            if (_fileSystem.FileExists(path))
            {
                _log.Debug("Deleting file " + path);
                _fileSystem.DeleteFile(path);
            }
        }

        /// <summary>
        /// The name of the project being compiled
        /// </summary>
        [Required]
        public string ProjectName { get; set; }

        /// <summary>
        /// The list of files in the project marked as Content or None
        /// </summary>
        [Required]
        public ITaskItem[] FilesInProject { get; set; }

        /// <summary>
        /// The directory in which the built files were written to
        /// </summary>
        [Required]
        public string OutDirectory { get; set; }

        /// <summary>
        /// The 32-bit .xll file path; set to <code>$(MSBuildThisFileDirectory)\ExcelDna.xll</code> by default
        /// </summary>
        [Required]
        public string Xll32FilePath { get; set; }

        /// <summary>
        /// The 64-bit .xll file path; set to <code>$(MSBuildThisFileDirectory)\ExcelDna64.xll</code> by default
        /// </summary>
        [Required]
        public string Xll64FilePath { get; set; }

        /// <summary>
        /// The name suffix for 32-bit .dna files
        /// </summary>
        public string FileSuffix32Bit { get; set; }

        /// <summary>
        /// The name suffix for 64-bit .dna files
        /// </summary>
        public string FileSuffix64Bit { get; set; }

        /// <summary>
        /// Enable/disable to have an .xll file with no packed assemblies
        /// </summary>
        public bool UnpackIsEnabled { get; set; }

        /// <summary>
        /// Enable/disable running ExcelDnaPack for .dna files
        /// </summary>
        public bool PackIsEnabled { get; set; }

        /// <summary>
        /// Packed add-in name suffix
        /// </summary>
        public string PackedFileSuffix { get; set; }

        /// <summary>
        /// Explicit 32-bit output file name
        /// </summary>
        public string Packed32BitXllName { get; set; }

        /// <summary>
        /// Explicit 64-bit output file name
        /// </summary>
        public string Packed64BitXllName { get; set; }

        /// <summary>
        /// Custom add-in file name
        /// </summary>
        public string AddInFileName { get; set; }

        /// <summary>
        /// The output directory for the 'published' add-in
        /// </summary>
        public string PublishPath { get; set; }
    }
}
