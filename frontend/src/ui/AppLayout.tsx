import { BarChartOutlined, LogoutOutlined, TransactionOutlined } from '@ant-design/icons';
import { Button, Layout, Menu, Typography } from 'antd';
import { Outlet, useLocation, useNavigate } from 'react-router-dom';

import { useAuthStore } from '../stores/authStore';

const { Header, Sider, Content } = Layout;

export function AppLayout() {
  const navigate = useNavigate();
  const location = useLocation();
  const logout = useAuthStore((state) => state.logout);
  const selectedKey = location.pathname.startsWith('/sales-payment') ? 'sales-payment' : 'dashboard';

  return (
    <Layout className="app-shell">
      <Sider width={248} className="app-sidebar">
        <Typography.Title level={4} className="app-logo">
          LMB 영업 관리
        </Typography.Title>
        <Menu
          mode="inline"
          selectedKeys={[selectedKey]}
          onClick={({ key }) => {
            if (key === 'dashboard') {
              navigate('/');
            }
            if (key === 'sales-payment') {
              navigate('/sales-payment');
            }
          }}
          items={[
            { key: 'dashboard', icon: <BarChartOutlined />, label: '대시보드' },
            { key: 'sales-payment', icon: <TransactionOutlined />, label: '영업 입금 자동화' },
          ]}
        />
      </Sider>
      <Layout>
        <Header className="app-header">
          <Typography.Text strong>SMB 영업/입금 자동화 MVP</Typography.Text>
          <Button
            icon={<LogoutOutlined />}
            onClick={() => {
              logout();
              navigate('/login');
            }}
          >
            로그아웃
          </Button>
        </Header>
        <Content className="app-content">
          <Outlet />
        </Content>
      </Layout>
    </Layout>
  );
}
