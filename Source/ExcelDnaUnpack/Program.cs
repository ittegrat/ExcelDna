﻿using System;
using System.Reflection;

//** TODO

//** MAYBE
// - overwrite make a clean extract ?
// - beautify resource names (problem: packing flattens tree and discards info)
// - Extract some types of unknown resources (version, strings, ...)

namespace ExcelDnaUnpack
{
  class Program
  {

    enum Command { None, Clean, Extract, List, Help }

    static void Main(string[] args) {

      try {

        string fileName = null;
        string outPath = null;
        Command command = Command.None;
        ResourceType rTypes = ResourceType.NONE;
        bool overwrite = false;
        bool mksubfolders = false;

        for (var i = 0; i < args.Length; ++i) {

          if (args[i][0] != '-') {
            fileName = args[i];
            continue;
          }

          switch (args[i].Substring(0, 2)) {

            case "--":
              switch (args[i].Substring(2)) {
                case "help":
                  command = Command.Help;
                  break;
                case "list-all":
                  rTypes = ResourceType.ALL;
                  command = Command.List;
                  break;
                default:
                  throw new ArgumentException($"Invalid option: {args[i]}");
              }
              break;

            case "-c":
              command = Command.Clean;
              break;

            case "-f":
              overwrite = true;
              break;

            case "-h":
              command = Command.Help;
              break;

            case "-l":
            case "-x":
              rTypes = ResourceType.KNOWN;
              if (args[i].Length > 2)
                rTypes = ParseCodes(args[i].Substring(2));
              command = args[i][1] == 'x' ? Command.Extract : Command.List;
              break;

            case "-o":
              ++i; outPath = args[i];
              break;

            case "-s":
              mksubfolders = true;
              break;

            default:
              throw new ArgumentException($"Invalid option: {args[i]}");

          }

        }

        if (fileName == null || command == Command.Help) {
          Usage();
          return;
        }

        if (command == Command.None) {
          command = Command.List;
          rTypes = ResourceType.KNOWN;
        }

        using (var module = new Module(fileName)) {
          if (command == Command.List)
            module.List(rTypes);
          else if (command == Command.Clean)
            module.Clean(outPath, overwrite);
          else
            module.Extract(rTypes, outPath, overwrite, mksubfolders);
        }

      }
      catch (Exception ex) {
        Environment.ExitCode = 1;
        Console.WriteLine(ex.Message);
      }

#if DEBUG
      Console.WriteLine("Press any key...");
      Console.ReadKey();
#endif

    }
    static ResourceType ParseCodes(string codes) {

      ResourceType ans = ResourceType.NONE;

      foreach (var c in codes) {
        switch (c) {
          case 'a':
            ans |= ResourceType.ASSEMBLY;
            break;
          case 'c':
            ans |= ResourceType.CONFIG;
            break;
          case 'd':
            ans |= ResourceType.DNA;
            break;
          case 'i':
            ans |= ResourceType.IMAGE;
            break;
          case 'p':
            ans |= ResourceType.PDB;
            break;
          case 's':
            ans |= ResourceType.SOURCE;
            break;
          case 't':
            ans |= ResourceType.TYPELIB;
            break;
          case 'u':
            ans |= ResourceType.UNKNOWN;
            break;
          default:
            throw new ArgumentException($"Invalid resource code: {c}");
        }
      }

      return ans;

    }
    static void Usage() {

      var name = System.IO.Path.GetFileNameWithoutExtension(Environment.GetCommandLineArgs()[0]);
      var arch = Environment.Is64BitProcess ? "64" : "32";

      var attributes = typeof(Program).Assembly.GetCustomAttributes(false);

      var version = String.Empty;
      var copyright = String.Empty;

      foreach (var a in attributes) {
        switch (a.GetType().Name) {
          case "AssemblyFileVersionAttribute":
            version = ((AssemblyFileVersionAttribute)a).Version;
            continue;
          case "AssemblyCopyrightAttribute":
            copyright = ((AssemblyCopyrightAttribute)a).Copyright;
            continue;
        }
      }

      string logo = String.Format(
        "{0} ({1} bit) v{2}\n" +
        "{3}\n" +
        "{4} ({5} bit)",
        name,
        arch,
        version,
        copyright,
        Environment.OSVersion,
        Environment.Is64BitOperatingSystem ? "64" : "32"
      );

      Console.Write(@"
{0}

Usage: {1} [options...] file

Options:
  -c: clean all known resources
  -f: overwrite existing files
  -h, --help: print this help
  -l[.]: List resources, where [.] is one or more resource codes
  --list-all: List all resources
  -o DIRECTORY: output directory
  -s: create subfolders
  -x[.]: eXtract resources, where [.] is one or more resource codes

Resource codes:
  a: assemblies
  c: config
  d: dna
  i: images
  p: pdbs
  s: sources
  t: typelibs
  u: unknown

", logo, name);

    }

  }
}
