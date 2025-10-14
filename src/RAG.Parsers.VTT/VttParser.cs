using RAG.Parsers.VTT.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace RAG.Parsers.VTT;

public class VttParser : IDisposable
{
    public ExtractOutput ToMarkdown(Stream data, ExtractOptions options = null)
    {
        var output = new StringBuilder();
        var result = new ExtractOutput
        {
            Output = string.Empty
        };

        using var reader = new StreamReader(data, Encoding.UTF8, true, 1024, leaveOpen: true);
        string line;
        string currentSpeaker = null;
        var cueText = new StringBuilder();
        bool inCue = false;
        string cueStart = null;
        string cueEnd = null;
        bool firstCue = true;

        // Regex for cue timing and optional identifier
        var cueHeaderRegex = new Regex(@"^(?<id>[a-fA-F0-9\-]{36}|[^\s]+)?\s*(?<start>\d{2}:\d{2}:\d{2}\.\d{3}) --> (?<end>\d{2}:\d{2}:\d{2}\.\d{3})", RegexOptions.Compiled);
        var speakerRegex = new Regex(@"<v ([^>]+)>", RegexOptions.Compiled);

        while ((line = reader.ReadLine()) != null)
        {
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("WEBVTT"))
            {
                continue;
            }

            var cueHeaderMatch = cueHeaderRegex.Match(line);
            if (cueHeaderMatch.Success)
            {
                if (inCue && cueText.Length > 0)
                {
                    if (!firstCue)
                    {
                        output.AppendLine(); // Ligne vide entre deux messages
                    }
                    else
                    {
                        firstCue = false;
                    }
                    // Ajout du timestamp et de la durée
                    if (!string.IsNullOrEmpty(cueStart) && !string.IsNullOrEmpty(cueEnd))
                    {
                        var start = TimeSpan.Parse(cueStart);
                        var end = TimeSpan.Parse(cueEnd);
                        var duration = end - start;
                        output.AppendLine($@"> [{cueStart} / {duration:hh\:mm\:ss\.fff}]" + Environment.NewLine);
                    }
                    output.AppendLine(cueText.ToString().Trim());
                    cueText.Clear();
                }
                inCue = true;
                // On ignore le GUID/id, on ne l'utilise pas dans le rendu
                cueStart = cueHeaderMatch.Groups["start"].Value;
                cueEnd = cueHeaderMatch.Groups["end"].Value;
                continue;
            }

            if (inCue)
            {
                // Detect speaker
                var speakerMatch = speakerRegex.Match(line);
                if (speakerMatch.Success)
                {
                    currentSpeaker = speakerMatch.Groups[1].Value;
                    var text = line.Substring(speakerMatch.Index + speakerMatch.Length).Trim();
                    cueText.AppendLine($"**{currentSpeaker}:** {text}");
                }
                else
                {
                    cueText.AppendLine(line.Trim());
                }
            }
        }
        // Flush last cue
        if (cueText.Length > 0)
        {
            if (!firstCue)
            {
                output.AppendLine();
            }
            if (!string.IsNullOrEmpty(cueStart) && !string.IsNullOrEmpty(cueEnd))
            {
                var start = TimeSpan.Parse(cueStart);
                var end = TimeSpan.Parse(cueEnd);
                var duration = end - start;
                output.AppendLine($@"> [{cueStart} / {duration:hh\:mm\:ss\.fff}]" + Environment.NewLine);
            }
            output.AppendLine(cueText.ToString().Trim());
        }

        result.Output = output.ToString().Trim();
        return result;
    }

    public void Dispose()
    {
        // Nothing to do here
    }
}
