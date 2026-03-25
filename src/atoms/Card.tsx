import React from 'react';

interface CardProps {
  children: React.ReactNode;
  className?: string;
  onClick?: () => void;
  hover?: boolean;
}

export const Card: React.FC<CardProps> = ({ 
  children, 
  className = '', 
  onClick, 
  hover = false 
}) => {
  const baseClasses = 'bg-white rounded-lg shadow-md border border-gray-200 transition-all duration-200';
  const hoverClasses = hover ? 'hover:shadow-lg hover:border-red-200 cursor-pointer' : '';
  const clickClasses = onClick ? 'hover:shadow-lg cursor-pointer' : '';
  
  const classes = [baseClasses, hoverClasses, clickClasses, className]
    .filter(Boolean)
    .join(' ');

  return (
    <div className={classes} onClick={onClick}>
      {children}
    </div>
  );
};