from typing import Optional
import json
import requests

class LLMResponse:
    def __init__(self):
        self.current_session_id = None

    def _handle_response(self, response: requests.Response) -> Optional[str]:
        try:
            # 尝试解析JSON响应
            try:
                data = response.json()
            except json.JSONDecodeError:
                print(f"\nError: Response is not valid JSON.")
                print(f"Raw Response: {response.text}")
                return None

            # 添加更详细的响应数据检查
            if not isinstance(data, dict):
                print(f"\nError: Unexpected response format - expected dictionary, got {type(data)}")
                print(f"Response: {data}")
                return None

            if 'data' not in data:
                print(f"\nError: Response missing 'data' field")
                print(f"Response: {data}")
                return None

            data_content = data.get('data', {})
            if not isinstance(data_content, dict):
                print(f"\nError: 'data' field is not a dictionary")
                print(f"Data content: {data_content}")
                return None

            answer = data_content.get('answer')
            if answer is not None:
                print(f"Assistant: {answer}")
                # 更新session_id
                self.current_session_id = data_content.get('session_id', self.current_session_id)
            else:
                print("\nError: No answer found in response")
                print(f"Data content: {data_content}")

        except Exception as e:
            print(f"\nError processing response: {str(e)}")
            print(f"Response content: {response.text}")
        return None 