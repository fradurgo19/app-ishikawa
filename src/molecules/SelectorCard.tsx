import React from 'react';
import { Card } from '../atoms/Card';
import { type LucideIcon } from 'lucide-react';

interface SelectorCardProps {
  title: string;
  description: string;
  icon: LucideIcon;
  onClick: () => void;
  selected?: boolean;
}

export const SelectorCard: React.FC<SelectorCardProps> = ({
  title,
  description,
  icon: Icon,
  onClick,
  selected = false,
}) => {
  return (
    <Card
      className={`p-6 ${selected ? 'ring-2 ring-red-500 bg-red-50' : ''}`}
      onClick={onClick}
      hover
    >
      <div className="text-center">
        <div className={`inline-flex p-4 rounded-full ${selected ? 'bg-red-100 text-red-600' : 'bg-gray-100 text-gray-600'}`}>
          <Icon size={32} />
        </div>
        <h3 className="mt-4 text-xl font-semibold text-gray-900">{title}</h3>
        <p className="mt-2 text-gray-600">{description}</p>
      </div>
    </Card>
  );
};