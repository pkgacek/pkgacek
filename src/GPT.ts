var CHAT_GPT_API_KEY =
  PropertiesService.getScriptProperties().getProperty("CHAT_GPT_API_KEY");
const BASE_URL = "https://api.openai.com/v1/chat/completions";
type HttpMethod = "get" | "delete" | "patch" | "post" | "put";
/**
 * @name fetchData
 * @description
 * @param systemContent
 * @param userContent
 * @returns
 */
function fetchData(systemContent: string, userContent: string): string {
  try {
    const headers = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${CHAT_GPT_API_KEY}`,
    };

    const options = {
      headers,
      method: "get" as HttpMethod,
      muteHttpExceptions: true,
      payload: JSON.stringify({
        model: "gpt-3.5-turbo",
        messages: [
          {
            role: "system",
            content: systemContent,
          },
          {
            role: "user",
            content: userContent,
          },
        ],
        temperature: 0.7,
      }),
    };
    const response = JSON.parse(
      UrlFetchApp.fetch(BASE_URL, options) as any as string,
    );
    if (response.error) {
      return response.error.message;
    }
    return response.choices[0].message.content;
  } catch (e) {
    Logger.log({ e });
    return "An error occurred. Please check your formula or try again later.";
  }
}
/**
 * @name GPT
 * @param userInput
 * @param systemInput
 * @return
 * @customfunction
 */
function GPT(systemInput: string, userInput: string) {
  return Array.isArray(userInput)
    ? userInput.flat().map((text) => fetchData(systemInput, text))
    : fetchData(systemInput, userInput);
}
