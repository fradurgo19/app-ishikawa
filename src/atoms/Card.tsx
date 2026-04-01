import React from 'react';

interface CardProps {
  children: React.ReactNode;
  className?: string;
  onClick?: () => void;
  hover?: boolean;
  /** Solo con onClick: el Card se renderiza como elemento button (teclado y lectores de pantalla). */
  'aria-pressed'?: boolean;
}

export const Card: React.FC<CardProps> = ({
  children,
  className = '',
  onClick,
  hover = false,
  'aria-pressed': ariaPressed,
}) => {
  const baseClasses = 'bg-white rounded-lg shadow-md border border-gray-200 transition-all duration-200';
  const hoverClasses = hover ? 'hover:shadow-lg hover:border-red-200 cursor-pointer' : '';
  const clickClasses = onClick ? 'hover:shadow-lg cursor-pointer' : '';

  const classes = [baseClasses, hoverClasses, clickClasses, className].filter(Boolean).join(' ');

  if (onClick) {
    return (
      <button
        type="button"
        className={`${classes} block w-full cursor-pointer text-left font-[inherit] antialiased focus:outline-none focus-visible:ring-2 focus-visible:ring-red-500 focus-visible:ring-offset-2`}
        onClick={onClick}
        aria-pressed={ariaPressed}
      >
        {children}
      </button>
    );
  }

  return <div className={classes}>{children}</div>;
};