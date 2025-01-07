/**
 * ChatGPT API를 호출하는 엑셀 사용자 함수
 * @customfunction
 * @param {string} question 사용자의 질문
 * @param {string} apiKey OpenAI API 키
 * @returns {Promise<string>} ChatGPT 응답
 */
async function GPT(question, apiKey) {
    if (!question) {
        return "❌ 질문을 입력하세요.";
    }
    if (!apiKey) {
        return "❌ API 키가 필요합니다.";
    }

    try {
        let response = await fetch("https://api.openai.com/v1/chat/completions", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: "gpt-4",
                messages: [{ role: "user", content: question }],
                temperature: 0.7,
                max_tokens: 150
            })
        });

        let data = await response.json();
        return data.choices[0].message.content;
    } catch (error) {
        return "❌ 오류 발생: " + error.message;
    }
}
