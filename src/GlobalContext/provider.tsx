import * as React from 'react';
import { useState } from 'react';
import UserContext from './context';
import { Footer } from './Footer';

interface  ProviderProps {
  children: React.ReactNode;
}

const Provider: React.FC<ProviderProps> = ({ children }) => {
  const [useHide, setHide] = useState(false);
  const [useFullscreen, setUseFullscreen] = useState(false);
  const [useId, setUseId] = useState(0);
  const [useActive, setuseActive] = useState(0);

  const toggleHide = () => setUseFullscreen(!useFullscreen);

  const toggleFullscreen = () => {
    if (document.fullscreenElement) {
      document.exitFullscreen();
    } else {
      document.documentElement.requestFullscreen();
    }
  };
 
  return (
    <UserContext.Provider value={{ useHide, setHide,useFullscreen, toggleHide, toggleFullscreen ,setUseId,useId,useActive,setuseActive }}>
      {children}
      <Footer/>
    </UserContext.Provider>
  );
};

export default Provider;