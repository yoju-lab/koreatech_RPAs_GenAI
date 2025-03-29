from openai import OpenAI

# client = OpenAI()
import os
from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()

# OpenAI API 호출
client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
)

completion = client.chat.completions.create(
    model="gpt-4o",
    messages=[
        {
            "role": "user",
            "content": "Write a short poem."
        }
    ]
)

print(completion)
# print(completion.choices[0].message.content)