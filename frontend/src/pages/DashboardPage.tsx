import { Space, Typography } from 'antd';

export function DashboardPage() {
  return (
    <Space className="dashboard-page" direction="vertical" size={18}>
      <div className="page-heading">
        <div>
          <Typography.Title level={3}>실적관리</Typography.Title>
        </div>
      </div>
    </Space>
  );
}
