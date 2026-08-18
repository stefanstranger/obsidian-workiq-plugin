import { describe, expect, it } from "vitest";
import {
  buildWorkIqSearchRequest,
  flattenWorkIqHits,
  formatWorkIqHits,
  WorkIqSearchResponse
} from "../src/workiq";

describe("WorkIQ search helpers", () => {
  it("builds a Microsoft Search request from a trimmed query", () => {
    expect(
      buildWorkIqSearchRequest("  project alpha  ", {
        entityTypes: ["driveItem", "message"],
        maxResults: 5
      })
    ).toEqual({
      requests: [
        {
          entityTypes: ["driveItem", "message"],
          from: 0,
          query: {
            queryString: "project alpha"
          },
          size: 5
        }
      ]
    });
  });

  it("rejects empty search queries", () => {
    expect(() =>
      buildWorkIqSearchRequest("   ", {
        entityTypes: ["driveItem"],
        maxResults: 10
      })
    ).toThrow("Enter a WorkIQ search query.");
  });

  it("flattens and formats Microsoft Search hits as note context", () => {
    const response: WorkIqSearchResponse = {
      value: [
        {
          hitsContainers: [
            {
              hits: [
                {
                  title: "Roadmap",
                  summary: "  Strategy   document  ",
                  url: "https://contoso.sharepoint.com/roadmap"
                },
                {
                  title: "Planning mail",
                  resource: {
                    summary: "Next planning steps",
                    webUrl: "https://outlook.office.com/mail/item"
                  }
                }
              ]
            }
          ]
        }
      ]
    };

    expect(formatWorkIqHits("roadmap", flattenWorkIqHits(response))).toBe(
      [
        "## WorkIQ search: roadmap",
        "",
        "1. [Roadmap](https://contoso.sharepoint.com/roadmap)",
        "   - Strategy document",
        "2. [Planning mail](https://outlook.office.com/mail/item)",
        "   - Next planning steps"
      ].join("\n")
    );
  });

  it("ignores unsafe result links and escapes Markdown text", () => {
    expect(
      formatWorkIqHits("security", [
        {
          title: "Result [one]",
          summary: "Summary with [brackets]",
          url: "http://contoso.example/insecure"
        }
      ])
    ).toBe(
      [
        "## WorkIQ search: security",
        "",
        "1. Result \\[one\\]",
        "   - Summary with \\[brackets\\]"
      ].join("\n")
    );
  });
});
