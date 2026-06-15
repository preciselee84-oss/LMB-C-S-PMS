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
        <Typography.Title level={3}>AX-Wetak 360</Typography.Title>
        <Typography.Paragraph type="secondary">
          Webcash의 We와 360도 빈틈없는 관리를 담은 AX-Wetak 360입니다.
        </Typography.Paragraph>
        <Typography.Paragraph type="secondary">관리자가 발급한 ID와 비밀번호로 로그인하세요.</Typography.Paragraph>
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
