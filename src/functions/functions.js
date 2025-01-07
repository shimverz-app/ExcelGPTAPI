/**
 * Excel 사용자 지정 함수: OpenAI API 호출
 * @param {string} question - 사용자 질문
 * @param {string} apiKey - OpenAI API 키
 * @returns {Promise<string>} ChatGPT 응답
 * @customfunction
 */
async function CHATGPT_ASK(question, apiKey) {
    const url = "https://api.openai.com/v1/chat/completions";

    const payload = {
        model: "gpt-4",  // 사용할 모델 선택 (gpt-4 또는 gpt-3.5-turbo)
        messages: [{ role: "user", content: question }],
        max_tokens: 150,
        temperature: 0.7,
    };

    try {
        const response = await fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${apiKey}`
            },
            body: JSON.stringify(payload),
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        return data.choices[0]?.message?.content || "응답 없음";
    } catch (error) {
        console.error("Error fetching ChatGPT response:", error);
        return "오류 발생";
    }
}

// Office.js 준비 완료 후 실행
Office.onReady(() => {
    console.log("Excel Add-in Ready");
});
