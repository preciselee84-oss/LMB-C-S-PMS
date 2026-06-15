import { Button, Card, Form, Input, Typography, message } from 'antd';
import { useNavigate } from 'react-router-dom';

import { login } from '../api/auth';
import { useAuthStore } from '../stores/authStore';

type LoginForm = {
  username: string;
  password: string;
};

export function LoginPage() {
  const navigate = useNavigate();
  const setToken = useAuthStore((state) => state.setToken);
  const [messageApi, contextHolder] = message.useMessage();

  const handleFinish = async (values: LoginForm) => {
    try {
      const token = await login(values);
      setToken(token.access_token);
      navigate('/');
    } catch {
      messageApi.error('아이디 또는 비밀번호를 확인해주세요.');
    }
  };

  return (
    <main className="login-page">
      {contextHolder}
      <Card className="login-card">
        <Typography.Title level={3}>위탁사업장 관리</Typography.Title>
        <Form layout="vertical" onFinish={handleFinish}>
          <Form.Item name="username" label="아이디" rules={[{ required: true }]}>
            <Input autoComplete="username" />
          </Form.Item>
          <Form.Item name="password" label="비밀번호" rules={[{ required: true }]}>
            <Input.Password autoComplete="current-password" />
          </Form.Item>
          <Button type="primary" htmlType="submit" block>
            로그인
          </Button>
        </Form>
      </Card>
    </main>
  );
}
