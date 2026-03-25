import React from 'react';
import { Card } from '../atoms/Card';
import { type LucideIcon } from 'lucide-react';

interface KPICardProps {
  title: string;
  value: number;
  icon: LucideIcon;
  color?: 'primary' | 'secondary' | 'success' | 'warning';
}

export const KPICard: React.FC<KPICardProps> = ({ 
  title, 
  value, 
  icon: Icon, 
  color = 'primary' 
}) => {
  const colorClasses = {
    primary: 'text-red-600 bg-red-50',
    secondary: 'text-gray-600 bg-gray-50',
    success: 'text-green-600 bg-green-50',
    warning: 'text-yellow-600 bg-yellow-50',
  };

  return (
    <Card className="p-6" hover>
      <div className="flex items-center">
        <div className={`p-3 rounded-lg ${colorClasses[color]}`}>
          <Icon size={24} />
        </div>
        <div className="ml-4">
          <p className="text-sm font-medium text-gray-600">{title}</p>
          <p className="text-3xl font-bold text-gray-900">{value}</p>
        </div>
      </div>
    </Card>
  );
};