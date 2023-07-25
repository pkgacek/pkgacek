const FOLDER_ID =
  PropertiesService.getScriptProperties().getProperty("FOLDER_ID");

/**
 * @name createResume
 * @description This generates a new resume as well as new cover letter file.
 * @returns
 */
function createResume(): {
  resumeId: string;
  coverLetterId: string;
  folderId: string;
} {
  const { resume, cover_letter, folderId, folder } = createFiles();
  const {
    id: resumeId,
    doc: resumeDoc,
    body: resumeBody,
    markdown: resumeMarkdown,
  } = resume;
  const {
    id: coverLetterId,
    doc: coverLetterDoc,
    body: coverLetterBody,
    markdown: coverLetterMarkdown,
  } = cover_letter;
  let resumeCurrent = resume.current;
  let coverLetterCurrent = cover_letter.current;
  const randomId = makeId();
  const currentDate = formatDate(new Date());
  const data = getAllSheetsData();
  const parsedData = Object.entries(data);
  DriveApp.getFolderById(folderId).setName(
    `Resumer - ${currentDate} - ${randomId}`,
  );
  if (FOLDER_ID) {
    var destination = DriveApp.getFolderById(FOLDER_ID);
    destination.addFolder(folder);
    folder.getParents().next().removeFolder(folder);
  }
  /**
   * @name setName
   * @description
   * @param name
   * @param doc
   * @param body
   * @param markdown
   * @param isCoverLetter
   * @returns
   */
  function setName(
    name: string,
    doc = resumeDoc,
    body = resumeBody,
    markdown = resumeMarkdown,
    isCoverLetter = false,
  ): {
    H1: GoogleAppsScript.Document.Paragraph;
    fileName: string;
  } {
    const H1 = setParagraph(body, name);
    H1.setAttributes({
      [DocumentApp.Attribute.HEADING]: DocumentApp.ParagraphHeading.HEADING1,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_SIZE]: 18,
    });
    markdown.push(`# **${name}**`);
    addLineBreak(H1, 3);
    H1.setAttributes({
      [DocumentApp.Attribute.BOLD]: true,
    });
    const fileName = `${
      isCoverLetter ? "Cover Letter" : "Resume"
    } - ${name} - ${currentDate} - ${randomId}`;
    doc.setName(fileName);
    return { H1, fileName };
  }
  /**
   * @name setHeadline
   * @description
   * @param headline
   * @param body
   * @param markdown
   * @param current
   * @returns
   */
  function setHeadline(
    headline: string,
    body = resumeBody,
    markdown = resumeMarkdown,
  ): GoogleAppsScript.Document.Paragraph {
    const P = setParagraph(body, headline);
    P.setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: 12,
      [DocumentApp.Attribute.BOLD]: true,
    });
    markdown.push("", `## **${headline}**`);
    addLineBreak(P, 3);
    return P;
  }
  /**
   * @name generateHeading
   * @description
   * @param value
   * @param body
   * @returns
   */
  function generateHeading(
    value: string,
    body = resumeBody,
    markdown = resumeMarkdown,
  ) {
    const H2 = body.appendParagraph(value);
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
   * @param url
   * @param config
   * @param config.indent
   * @param config.fontSize
   * @param config.bold
   * @param config.italic
   * @param markdown
   * @param current
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
    body = resumeBody,
    markdown = resumeMarkdown,
  ): GoogleAppsScript.Document.Paragraph {
    const indentValue = config && config.indent ? MARKDOWN_INDENT : "";
    let current: GoogleAppsScript.Document.Paragraph;
    if (url) {
      current = setParagraph(body, `${key}: `, {
        indent: config?.indent || 0,
      });
      const paragraphUrl = setParagraph(body, value, {
        indent: config?.indent || 0,
      }).setLinkUrl(url);
      paragraphUrl.merge();
      markdown.push(`${indentValue}- ${key}: [${value}](${url})`);
    } else {
      current = setParagraph(body, `${key}: ${value}`, {
        indent: config?.indent || 0,
      });
      markdown.push(`${indentValue}- ${key}: ${value}`);
    }
    current.setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: config?.fontSize || 9,
      [DocumentApp.Attribute.ITALIC]: config?.italic || true,
      [DocumentApp.Attribute.BOLD]: config?.bold || false,
    });

    return current;
  }
  /**
   * @name generateEmptyParagraph
   * @description
   * @param config
   * @param config.indent
   * @param current
   * @returns
   */
  function generateEmptyParagraph(
    config?: { indent?: number },
    current = resumeCurrent,
  ) {
    current = setParagraph(resumeBody, "", { indent: config?.indent || 1 });
    current.setAttributes({
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
   * @param config.indent
   * @param markdown
   * @returns
   */
  function generateInformations(
    key: string,
    value: string,
    config?: {
      indent?: number;
    },
    markdown = resumeMarkdown,
  ) {
    const current = setParagraph(resumeBody, key, {
      indent: config?.indent || 1,
    });
    current.setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: 9,
      [DocumentApp.Attribute.ITALIC]: false,
      [DocumentApp.Attribute.BOLD]: true,
    });
    current.appendText(`: ${value}`).setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: 9,
      [DocumentApp.Attribute.ITALIC]: false,
      [DocumentApp.Attribute.BOLD]: false,
    });
    addLineBreak(current, 3);
    markdown.push(
      `${config?.indent ? MARKDOWN_INDENT : ""}- **${key}**: ${value}`,
    );

    return current;
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
          const LI = setListItem(resumeBody, "", 0);
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
          resumeMarkdown.push(titleMarkdown);
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
            resumeCurrent = generateDetail(
              getFormattedDatePrefix(row),
              formattedDate,
              "",
              {
                indent,
              },
            );
          if (formattedLocation)
            resumeCurrent = generateDetail("Location", formattedLocation, "", {
              indent,
            });
          if (grade)
            resumeCurrent = generateDetail("Grade", grade, "", { indent });
          if (thesis)
            resumeCurrent = generateDetail("Thesis", thesis, "", { indent });
          if (cause)
            resumeCurrent = generateDetail("Cause", cause, "", { indent });
          if (credentialId || credentialUrl)
            resumeCurrent = generateDetail(
              "Credential",
              `${credentialId || "Check credential"}`,
              credentialUrl,
              { indent },
            );
          addLineBreak(resumeCurrent, 3);
          if (description) {
            const DESCRIPTION_LIST_ITEMS = description
              .split("\n")
              .map((el, idx, arr) => {
                const nestingLevel = calcNestingLevel(el) + 1;
                const formattedString = el.replace(/^((\s+-)|-) /, "");
                const ListItem = setListItem(
                  resumeBody,
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
                resumeMarkdown.push(
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
              resumeCurrent = generateInformations(
                "Technologies used",
                technologies,
                {
                  indent,
                },
              );
            if (concepts)
              resumeCurrent = generateInformations("Concepts used", concepts, {
                indent,
              });
            if (links) {
              const LINK_PARAGRAPH = setParagraph(resumeBody, "Links: ", {
                indent,
              });
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
              resumeMarkdown.push(markdownLinkList);
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
        const { fileName: resumeFileName, H1: resumeH1 } = setName(name);
        resumeCurrent = resumeH1;
        const { fileName: coverLetterFileName, H1: coverLetterH1 } = setName(
          name,
          coverLetterDoc,
          coverLetterBody,
          coverLetterMarkdown,
          true,
        );
        coverLetterCurrent = coverLetterH1;
      }
      if (headline) {
        resumeCurrent = setHeadline(headline);
        coverLetterCurrent = setHeadline(
          headline,
          coverLetterBody,
          coverLetterMarkdown,
        );
      }
      if (address) {
        resumeCurrent = generateDetail("Address", address);
        coverLetterCurrent = generateDetail(
          "Address",
          address,
          undefined,
          undefined,
          coverLetterBody,
          coverLetterMarkdown,
        );
      }
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
          resumeCurrent = generateDetail(label, text, link);
          coverLetterCurrent = generateDetail(
            label,
            text,
            link,
            undefined,
            coverLetterBody,
            coverLetterMarkdown,
          );
          Logger.log({ coverLetterCurrent });
        });
        addLineBreak(resumeCurrent, 3);
        addLineBreak(coverLetterCurrent, 12);
      }
      if (technologies || concepts || languages || traits) {
        let HEADING = generateHeading(k);
        if (technologies)
          resumeCurrent = generateInformations("Technologies", technologies);
        if (concepts)
          resumeCurrent = generateInformations("Concepts", concepts);
        if (traits) resumeCurrent = generateInformations("Traits", traits);
        if (languages)
          resumeCurrent = generateInformations("Languages", languages);
        HEADING.setAttributes({ [DocumentApp.Attribute.BOLD]: true });
      }
    }
  }
  addLineBreak(resumeCurrent, 15);
  resumeCurrent = setParagraph(resumeBody, CONSENT);
  resumeCurrent.setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 8,
    [DocumentApp.Attribute.BOLD]: false,
    [DocumentApp.Attribute.ITALIC]: true,
  });
  resumeMarkdown.push("\n", `*${CONSENT}*`);

  const coverLetterValue = getCoverLetter();
  coverLetterCurrent = setParagraph(
    coverLetterBody,
    coverLetterValue,
  ).setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 12,
    [DocumentApp.Attribute.BOLD]: false,
    [DocumentApp.Attribute.ITALIC]: false,
  });

  coverLetterMarkdown.push("\n", coverLetterValue);

  const formattedMarkdown = resumeMarkdown.join("\n");
  const formattedCoverLetterMarkdown = coverLetterMarkdown.join("\n");

  addResume(formattedMarkdown, formattedCoverLetterMarkdown);

  removeEmptyParagraph(resumeBody);
  removeEmptyParagraph(coverLetterBody);

  resumeBody.setAttributes({ [DocumentApp.Attribute.FONT_FAMILY]: "Lato" });
  coverLetterBody.setAttributes({
    [DocumentApp.Attribute.FONT_FAMILY]: "Lato",
  });

  resumeDoc.saveAndClose();
  coverLetterDoc.saveAndClose();

  Logger.log(
    `Finished
      - check resume at https://docs.google.com/document/d/${resumeId}
      - check cover letter at https://docs.google.com/document/d/${coverLetterId}
      - check folder at https://drive.google.com/drive/folders/${folderId}
      `,
  );
  return { resumeId, coverLetterId, folderId };
}
