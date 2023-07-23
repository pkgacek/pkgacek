/**
 * @name createResume
 * @description
 * @returns
 */
function createResume() {
  let markdown = [];
  var CURRENT:
    | GoogleAppsScript.Document.Paragraph
    | GoogleAppsScript.Document.ListItem;
  const DOC = DocumentApp.create("Resumer");
  const DOC_ID = DOC.getId();
  const BODY = DOC.getBody();
  addResume("");
  BODY.setText("");
  deleteAllParagraphs(BODY);
  BODY.clear();
  const data = getAllSheetsData();
  const parsedData = Object.entries(data);
  /**
   * @name generateHeading
   * @description
   * @param value
   * @returns
   */
  function generateHeading(value: string) {
    const H2 = BODY.appendParagraph(value);
    H2.setAttributes({
      [DocumentApp.Attribute.BOLD]: true,
    });
    H2.setAttributes({
      [DocumentApp.Attribute.HEADING]: DocumentApp.ParagraphHeading.HEADING2,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_SIZE]: 14,
    });
    addLineBreak(H2, 3);
    markdown.push("", `### ${value}`, "---");
    return H2;
  }
  /**
   * @name generateDetail
   * @description
   * @param key
   * @param value
   * @param indent
   * @param url
   * @returns
   */
  function generateDetail(
    key: string,
    value: string,
    url?: string,
    config?: {
      indent?: number;
      fontSize?: number;
      bold?: boolean;
      italic?: boolean;
    },
  ) {
    const indentValue = config && config.indent ? MARKDOWN_INDENT : "";
    if (url) {
      CURRENT = setParagraph(BODY, `${key}: `, {
        indent: config?.indent || 0,
      });
      const paragraphUrl = setParagraph(BODY, value, {
        indent: config?.indent || 0,
      }).setLinkUrl(url);
      paragraphUrl.merge();
      markdown.push(`${indentValue}- ${key}: [${value}](${url})`);
    } else {
      CURRENT = setParagraph(BODY, `${key}: ${value}`, {
        indent: config?.indent || 0,
      });
      markdown.push(`${indentValue}- ${key}: ${value}`);
    }
    CURRENT.setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: config?.fontSize || 9,
      [DocumentApp.Attribute.ITALIC]: config?.italic || true,
      [DocumentApp.Attribute.BOLD]: config?.bold || false,
    });
  }
  /**
   * @name generateEmptyParagraph
   * @description
   * @param config
   * @returns
   */
  function generateEmptyParagraph(config?: { indent?: number }) {
    CURRENT = setParagraph(BODY, "", { indent: config?.indent || 1 });
    CURRENT.setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: 0,
      [DocumentApp.Attribute.ITALIC]: false,
      [DocumentApp.Attribute.BOLD]: false,
    });
  }
  /**
   * @name generateInformations
   * @description
   * @param key
   * @param value
   * @param config
   * @returns
   */
  function generateInformations(
    key: string,
    value: string,
    config?: {
      indent?: number;
    },
  ) {
    CURRENT = setParagraph(BODY, key, { indent: config?.indent || 1 });
    CURRENT.setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: 9,
      [DocumentApp.Attribute.ITALIC]: false,
      [DocumentApp.Attribute.BOLD]: true,
    });
    CURRENT.appendText(`: ${value}`).setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: 9,
      [DocumentApp.Attribute.ITALIC]: false,
      [DocumentApp.Attribute.BOLD]: false,
    });
    addLineBreak(CURRENT, 3);
    markdown.push(
      `${config?.indent ? MARKDOWN_INDENT : ""}- **${key}**: ${value}`,
    );
  }
  for (let i = 0; i < parsedData.length; i++) {
    const [k, v] = parsedData[i];
    if (isSection(v)) {
      let HEADING = generateHeading(k);
      for (let j = 0; j < v.length; j++) {
        const row = v[j] as {
          enable: boolean;
          url: string;
          organization: string;
          location: string;
          "location type": string;
          description: string;
          technologies: string;
          concepts: string;
          grade: string;
          thesis: string;
          links: string;
          cause: string;
          level: string;
          "credential id": string;
          "credential url": string;
          name: string;
          headline: string;
          languages: string;
        };
        if (row.enable) {
          const { url, organization } = row;
          const title = getFormattedTitle(row);
          const LI = setListItem(BODY, "", 0);
          LI.appendText(title);
          let titleMarkdown = `- #### **${title}**`;
          if (organization) {
            const formattedAt = getFormattedAt(row);
            LI.appendText(` ${formattedAt} `);
            const organization = LI.appendText(row.organization);
            if (url) {
              organization.setLinkUrl(url);
              titleMarkdown += ` **${formattedAt} [${row.organization}](${url})**`;
            } else {
              titleMarkdown += ` **${formattedAt} ${row.organization}**`;
            }
          }
          LI.setAttributes({
            [DocumentApp.Attribute.HEADING]:
              DocumentApp.ParagraphHeading.HEADING3,
            [DocumentApp.Attribute.FONT_SIZE]: 12,
            [DocumentApp.Attribute.BOLD]: true,
            [DocumentApp.Attribute.ITALIC]: false,
          });
          addLineBreak(LI, 3);
          markdown.push(titleMarkdown);
          const indent: number = LI.getIndentStart();
          const {
            location,
            "location type": locationType,
            description,
            technologies,
            concepts,
            grade,
            thesis,
            links,
            cause,
            "credential id": credentialId,
            "credential url": credentialUrl,
          } = row;
          const formattedDate = getFormattedDate(row);
          const formattedLocation = getFormattedLocation(
            location,
            locationType,
          );
          if (formattedDate)
            generateDetail(getFormattedDatePrefix(row), formattedDate, "", {
              indent,
            });
          if (formattedLocation)
            generateDetail("Location", formattedLocation, "", { indent });
          if (grade) generateDetail("Grade", grade, "", { indent });
          if (thesis) generateDetail("Thesis", thesis, "", { indent });
          if (cause) generateDetail("Cause", cause, "", { indent });
          if (credentialId || credentialUrl)
            generateDetail(
              "Credential",
              `${credentialId || "Check credential"}`,
              credentialUrl,
              { indent },
            );
          addLineBreak(CURRENT, 3);
          if (description) {
            const DESCRIPTION_LIST_ITEMS = description
              .split("\n")
              .map((el, idx, arr) => {
                const nestingLevel = calcNestingLevel(el) + 1;
                const formattedString = el.replace(/^((\s+-)|-) /, "");
                const ListItem = setListItem(
                  BODY,
                  formattedString,
                  nestingLevel,
                );
                ListItem.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
                ListItem.setAttributes({
                  [DocumentApp.Attribute.FONT_SIZE]: 9,
                  [DocumentApp.Attribute.ITALIC]: false,
                  [DocumentApp.Attribute.BOLD]: false,
                });
                if (technologies || concepts || links) {
                  addLineBreak(ListItem, idx === arr.length - 1 ? 3 : 1.5);
                } else if (idx !== arr.length - 1) {
                  addLineBreak(ListItem, 1.5);
                }
                markdown.push(
                  `${MARKDOWN_INDENT.repeat(
                    nestingLevel + 1,
                  )}- ${formattedString}`,
                );
                return ListItem;
              })
              .forEach((li) => {
                li.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
              });
          }
          if (!(technologies || concepts || links)) {
            generateEmptyParagraph({ indent });
          } else {
            if (technologies)
              generateInformations("Technologies used", technologies, {
                indent,
              });
            if (concepts)
              generateInformations("Concepts used", concepts, { indent });
            if (links) {
              const LINK_PARAGRAPH = setParagraph(BODY, "Links: ", { indent });
              LINK_PARAGRAPH.setAttributes({
                [DocumentApp.Attribute.FONT_SIZE]: 9,
                [DocumentApp.Attribute.ITALIC]: false,
                [DocumentApp.Attribute.BOLD]: true,
              });
              let markdownLinkList = `${MARKDOWN_INDENT}- **Links:** `;
              links.split("\n").forEach((el, idx, arr) => {
                const isLast = idx === arr.length - 1;
                const isSecondToLast = idx === arr.length - 2;
                const separator = isSecondToLast ? ", and " : ", ";
                const formattedString = el.replace(/^((\s+-)|-) /, "");
                const [text, link] = getLink(formattedString);
                LINK_PARAGRAPH.appendText(text)
                  .setLinkUrl(link)
                  .setAttributes({
                    [DocumentApp.Attribute.FONT_SIZE]: 9,
                    [DocumentApp.Attribute.ITALIC]: false,
                    [DocumentApp.Attribute.BOLD]: false,
                  });
                if (!isLast) {
                  LINK_PARAGRAPH.appendText(separator).setLinkUrl(null);
                }
                markdownLinkList += `${formattedString}${separator}`;
              });
              markdown.push(markdownLinkList);
              addLineBreak(LINK_PARAGRAPH, 3);
            }
          }
          LI.setGlyphType(DocumentApp.GlyphType.BULLET);
        }
      }
      HEADING.setAttributes({ [DocumentApp.Attribute.BOLD]: true });
    } else {
      const row = v[0];
      const {
        name,
        headline,
        address,
        links,
        languages,
        technologies,
        concepts,
        traits,
      } = row as {
        description: string;
        technologies: string;
        concepts: string;
        address: string;
        links: string;
        name: string;
        headline: string;
        languages: string;
        traits: string;
      };
      if (name) {
        const H1 = setParagraph(BODY, name);
        H1.setAttributes({
          [DocumentApp.Attribute.HEADING]:
            DocumentApp.ParagraphHeading.HEADING1,
          [DocumentApp.Attribute.BOLD]: true,
          [DocumentApp.Attribute.FONT_SIZE]: 18,
        });
        markdown.push(`# **${name}**`);
        CURRENT = H1;
        addLineBreak(CURRENT, 3);
        CURRENT.setAttributes({
          [DocumentApp.Attribute.BOLD]: true,
        });
        DOC.setName(`${name} - ${formatDate(new Date())} - ${makeId()}`);
      }
      if (headline) {
        const P = setParagraph(BODY, headline);
        P.setAttributes({
          [DocumentApp.Attribute.FONT_SIZE]: 12,
          [DocumentApp.Attribute.BOLD]: true,
        });
        markdown.push("", `## **${headline}**`);
        CURRENT = P;
        addLineBreak(CURRENT, 3);
      }
      if (address) generateDetail("Address", address);
      if (links) {
        links.split("\n").forEach((el) => {
          const formattedString = el.replace(/^((\s+-)|-) /, "");
          const [text, link] = getLink(formattedString);
          let label = "Link";
          if (link.includes("mailto:")) {
            label = "Mail";
          }
          if (link.includes("tel:")) {
            label = "Tel";
          }
          generateDetail(label, text, link);
        });
        addLineBreak(CURRENT, 3);
      }
      if (technologies || concepts || languages || traits) {
        let HEADING = generateHeading(k);
        if (technologies) generateInformations("Technologies", technologies);
        if (concepts) generateInformations("Concepts", concepts);
        if (traits) generateInformations("Traits", traits);
        if (languages) generateInformations("Languages", languages);
        HEADING.setAttributes({ [DocumentApp.Attribute.BOLD]: true });
      }
    }
  }
  addLineBreak(CURRENT, 15);
  const P = setParagraph(BODY, CONSENT);
  P.setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 8,
    [DocumentApp.Attribute.BOLD]: false,
    [DocumentApp.Attribute.ITALIC]: true,
  });
  markdown.push("\n", `*${CONSENT}*`);

  const formattedMarkdown = markdown.join("\n");
  addResume(formattedMarkdown);
  removeEmptyParagraph(BODY);
  BODY.setAttributes({ [DocumentApp.Attribute.FONT_FAMILY]: "Lato" });
  DOC.saveAndClose();
  Logger.log(
    `Finished - check document at https://docs.google.com/document/d/${DOC_ID}/edit`,
  );
  return `https://docs.google.com/document/d/${DOC_ID}/edit`;
}
